using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Audacia.Core;
using Audacia.Core.Extensions;
using Audacia.Spreadsheets.Extensions;
using Audacia.Spreadsheets.Validation;

namespace Audacia.Spreadsheets
{
    /// <summary>
    /// A default worksheet importer, can be overriden for extended logic.
    /// </summary>
    public class WorksheetImporter<TRowModel> where TRowModel : class, new()
    {
        private readonly string[] DateTimeFormats = new[]
        {
            "dd/MM/yyyy HH:mm:ss K",    // DateTimeOffset
            "dd/MM/yyyy HH:mm:ss",      // Long DateTime
            "dd/MM/yyyy HH:mm",         // Short DateTime
            "dd/MM/yyyy",               // Short Date
            "yyyy-MM-dd HH:mm:ss K",    // Sortable DateTimeOffset
            "yyyy-MM-dd HH:mm:ss",      // Sortable Long DateTime
            "yyyy-MM-dd HH:mm",         // Sortable Short DateTime
            "yyyy-MM-dd",               // Sortable Date
            "yyyy-MM-ddTH:mm:ss.fffK",  // ISO8601
            "O",                        // Round-trip DateTime
            "R",                        // RFC1123
            "u"                         // Universal Sortable DateTime
        };
        private bool skipWorksheetColumnMapping;

        /// <summary>
        /// Maps expected column headers to properties on the row model.
        /// </summary>
        protected IDictionary<string, PropertyInfo> ExpectedColumns { get; private set; } = new Dictionary<string, PropertyInfo>();

        /// <summary>
        /// Maps actual column headers to field index on the spreadsheet.
        /// </summary>
        protected IDictionary<string, int> SpreadsheetColumns { get; private set; } = new Dictionary<string, int>();

        /// <summary>
        /// Current row being parsed by the importer.
        /// </summary>
        protected TableRow CurrentRow { get; private set; }


        /// <summary>
        /// Manually map an expected column to a property on the row model.
        /// </summary>
        /// <param name="propertyExpression">Property on row model</param>
        public WorksheetImporter<TRowModel> MapColumn(Expression<Func<TRowModel, object>> propertyExpression)
        {
            var propertyInfo = ExpressionExtensions.GetPropertyInfo(propertyExpression);
            var columnHeader = propertyInfo.GetDataAnnotationDisplayName();

            return MapColumn(columnHeader, propertyExpression);
        }

        /// <summary>
        /// Manually map an expected column to a property on the row model.
        /// </summary>
        /// <param name="columnHeader">Expected column header or display name</param>
        /// <param name="propertyExpression">Property on row model</param>
        public WorksheetImporter<TRowModel> MapColumn(string columnHeader, Expression<Func<TRowModel, object>> propertyExpression)
        {
            if (ExpectedColumns.ContainsKey(columnHeader))
            {
                throw new InvalidOperationException($"Column '{columnHeader}' has already been mapped");
            }

            var propertyInfo = ExpressionExtensions.GetPropertyInfo(propertyExpression);
            ExpectedColumns.Add(columnHeader, propertyInfo);

            // Manually append spreadsheet cell mapping in the case where no column headers exist
            if (skipWorksheetColumnMapping)
            {
                var previousColumnIndex = SpreadsheetColumns.Any()
                    ? SpreadsheetColumns.Values.LastOrDefault()
                    : -1;
                SpreadsheetColumns.Add(columnHeader, previousColumnIndex + 1);
            }

            return this;
        }

        /// <summary>
        /// Parses the worksheet, is only able to handle data table style worksheets with column headers.
        /// If there are no column headers on the sheet, you will need to manually configure the expected columns to use this.
        /// </summary>
        /// <param name="worksheet">Worksheet to be parsed</param>
        /// <param name="ignoreProperties">Properties to ignore when generating expected column headers</param>
        public IEnumerable<ImportRow<TRowModel>> ParseWorksheet(WorksheetBase worksheet, params string[] ignoreProperties)
        {
            // We only support single worksheets
            var sheet = worksheet as Worksheet;
            if (sheet == null)
            {
                throw new InvalidCastException($"The worksheet being imported must inherit from {typeof(Worksheet).FullName}");
            }

            // Sets the expected column headers using the default column headers generated for the row model.
            if (!ExpectedColumns.Any())
            {
                ExpectedColumns = Tables
                    .GetColumns<TRowModel>(ignoreProperties)
                    .ToDictionary(tc => tc.Name, tc => tc.PropertyInfo);
            }

            // Create column headers map, if not manually setup
            if (!SpreadsheetColumns.Any())
            {
                if (skipWorksheetColumnMapping)
                {
                    throw new InvalidOperationException(
                        $"Incorrect usage, .{nameof(SkipColumnHeaderMapping)}() should be called before .{nameof(MapColumn)}() when no column headers are expected.");
                }

                // Check for duplicate column names in spreadsheet
                var duplicateColumnNames = sheet.Table.Columns
                    .Where(c => !string.IsNullOrWhiteSpace(c.Name))
                    .Where(c => sheet.Table.Columns.Count(tc => tc.Name == c.Name) > 1)
                    .Select(c => c.Name)
                    .ToArray();

                if (duplicateColumnNames.Any())
                {
                    yield return new ImportRow<TRowModel>
                    {
                        ImportErrors = new[] { new DuplicateColumnError(duplicateColumnNames) }
                    };
                    yield break;
                }

                // Convert column headers on spreadsheet into mapping dictionary
                SpreadsheetColumns = sheet.Table.Columns.ToDictionary();

                // Check for missing column headers
                var missingColumnNames = ExpectedColumns.Keys
                    .Where(expected => !SpreadsheetColumns.ContainsKey(expected))
                    .ToArray();

                if (missingColumnNames.Any())
                {
                    yield return new ImportRow<TRowModel>
                    {
                        ImportErrors = new[] { new MissingColumnError(missingColumnNames) }
                    };
                    yield break;
                }
            }

            // Iterate over and parse all rows
            foreach (var row in sheet.Table.Rows)
            {
                CurrentRow = row;
                var rowParseErrors = ParseRow(out var rowModel);

                // Allow for custom row validation if inherited
                var customValidationErrors = rowParseErrors.Any()
                    ? Enumerable.Empty<IImportError>()
                    : ValidateRow(rowModel);

                var importModel = new ImportRow<TRowModel>
                {
                    RowId = row.Id ?? 0,
                    Data = rowModel,
                    ImportErrors = rowParseErrors.Concat(customValidationErrors).ToArray()
                };

                // We're using yield return to allow for developers to design large imports where the memory can be garbage collected
                // Obviously this is a band aid against the raging typhoon that is the DocumentFormat.OpenXml library
                // Because of this design choice we can't have a global validation error list
                yield return importModel;
            }
        }

        /// <summary>
        /// Overrides automatic column header mapping when parsing the worksheet.
        /// This should be used in situations where a worksheet has no column headers.
        /// </summary>
        public WorksheetImporter<TRowModel> SkipColumnHeaderMapping()
        {
            skipWorksheetColumnMapping = true;
            return this;
        }

        /// <summary>
        /// Gets the column header for a property.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <exception cref="InvalidOperationException">When the property is not an expected worksheet column</exception>
        protected string GetColumnHeader(Expression<Func<TRowModel, object>> propertyExpression)
        {
            var propertyInfo = ExpressionExtensions.GetPropertyInfo(propertyExpression);

            if (!ExpectedColumns.Values.Contains(propertyInfo))
            {
                throw new InvalidOperationException($"Property '{propertyInfo.Name}' is not an expected worksheet column.");
            }

            return ExpectedColumns.Single(kvp => kvp.Value == propertyInfo).Key;
        }

        /// <summary>
        /// Returns the ID of the current row (defaults to zero).
        /// </summary>
        protected int GetRowNumber() 
        {
            // Row ID will always exist when parsing spreadsheets read from file, it won't exist if someone attempts to parse a worksheet generated for export
            return CurrentRow.Id ?? 0;
        }

        /// <summary>
        /// Handles the parsing of the CurrentRow, should add validation errors where necessary.
        /// </summary>
        protected virtual IEnumerable<IImportError> ParseRow(out TRowModel model)
        {
            // For the application developer to override
            // Use either TryGetCell(), TryGetValue(), TryGetValueStr()
            // or var fields = CurrentRow.GetFields() and GetValue(fields, ...)
            var rowErrors = new List<IImportError>();
            model = new TRowModel();

            // Iterate through all properties and set them based on cell value
            foreach (var expectedProperty in ExpectedColumns)
            {
                // Unable to find expected cell
                var columnName = expectedProperty.Key;
                if (!TryGetCell(columnName, out var cell))
                {
                    rowErrors.Add(new FieldMissingError(GetRowNumber(), columnName));
                    continue;
                }

                var classMember = expectedProperty.Value;
                var propertyType = classMember.PropertyType;
                var underlyingType = propertyType.GetUnderlyingTypeIfNullable();
                var isNullable = propertyType.IsNullable();

                // Get the cell value and parse it
                var parsedValue = ParseValue(columnName, underlyingType, isNullable, cell, rowErrors);

                // Set cell value
                if (classMember.CanWrite && (isNullable || parsedValue != null))
                {
                    classMember.SetValue(model, parsedValue);
                }
            }

            return rowErrors;
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="bool"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetBoolean(Expression<Func<TRowModel, object>> propertyExpression, out bool value)
        {
            value = false;
            return TryGetString(propertyExpression, out var str) &&
                   Bool.TryParse(str, out value);
        }

        /// <summary>
        /// Try get cell data from the current row.
        /// Cell.Value can be a string, decimal, or DateTime.
        /// Cell.FillColour is also parsed.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="cell">Cell Data</param>
        protected bool TryGetCell(Expression<Func<TRowModel, object>> propertyExpression, out TableCell value)
        {
            try
            {
                var column = GetColumnHeader(propertyExpression);
                return TryGetCell(column, out value);
            }
            catch
            {
                value = null;
                return false;
            }
        }

        /// <summary>
        /// Try get cell data from the current row.
        /// Cell.Value can be a string, decimal, or DateTime.
        /// Cell.FillColour is also parsed.
        /// </summary>
        /// <param name="columnHeader">Column Header</param>
        /// <param name="cell">Cell Data</param>
        protected bool TryGetCell(string columnHeader, out TableCell cell)
        {
            cell = null;

            if (!SpreadsheetColumns.ContainsKey(columnHeader)) return false;

            var columnIndex = SpreadsheetColumns[columnHeader];

            if (CurrentRow.Cells.Count <= columnIndex) return false;

            cell = CurrentRow.Cells[columnIndex];

            return true;
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="DateTime"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetDateTime(Expression<Func<TRowModel, object>> propertyExpression, out DateTime value)
        {
            // Further optimisation possible can be done but code makes code look overly complex
            // Could use TryGetCell(), then checking if the date was already parsed like in ParseValue()
            value = DateTime.MinValue;
            return TryGetString(propertyExpression, out var str) &&
                   TryParseDateTime(str, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="DateTimeOffset"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetDateTimeOffset(Expression<Func<TRowModel, object>> propertyExpression, out DateTimeOffset value)
        {
            // Further optimisation possible can be done but code makes code look overly complex
            // Could use TryGetCell(), then checking if the date was already parsed like in ParseValue()
            value = DateTimeOffset.MinValue;
            return TryGetString(propertyExpression, out var str) &&
                   TryParseDateTimeOffset(str, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="decimal"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetDecimal(Expression<Func<TRowModel, object>> propertyExpression, out decimal value)
        {
            // Further optimisation possible can be done but code makes code look overly complex
            // Could use TryGetCell(), then checking if the date was already parsed like in ParseValue()
            value = 0m;
            return TryGetString(propertyExpression, out var str) &&
                   decimal.TryParse(str, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="double"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetDouble(Expression<Func<TRowModel, object>> propertyExpression, out double value)
        {
            // Further optimisation possible can be done but code makes code look overly complex
            // Could use TryGetCell(), then checking if the date was already parsed like in ParseValue()
            value = 0d;
            return TryGetString(propertyExpression, out var str) &&
                   double.TryParse(str, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="{TEnum}"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetEnum<TEnum>(Expression<Func<TRowModel, object>> propertyExpression, out TEnum value)
            where TEnum : struct
        {
            // Further optimisation possible can be done but code makes code look overly complex
            // Could use TryGetCell(), then checking if the date was already parsed like in ParseValue()
            value = default(TEnum);
            return TryGetString(propertyExpression, out var str) &&
                   EnumMember.TryParse<TEnum>(str, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="float"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetFloat(Expression<Func<TRowModel, object>> propertyExpression, out float value)
        {
            // Further optimisation possible can be done but code makes code look overly complex
            // Could use TryGetCell(), then checking if the date was already parsed like in ParseValue()
            value = 0f;
            return TryGetString(propertyExpression, out var str) &&
                   float.TryParse(str, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="int"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetInteger(Expression<Func<TRowModel, object>> propertyExpression, out int value)
        {
            // Further optimisation possible can be done but code makes code look overly complex
            // Could use TryGetCell(), then checking if the date was already parsed like in ParseValue()
            value = 0;
            return TryGetString(propertyExpression, out var str) &&
                   int.TryParse(str, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="TimeSpan"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetTimespan(Expression<Func<TRowModel, object>> propertyExpression, out TimeSpan value)
        {
            // Further optimisation possible can be done but code makes code look overly complex
            // Could use TryGetCell(), then checking if the date was already parsed like in ParseValue()
            value = default(TimeSpan);
            return TryGetString(propertyExpression, out var str) &&
                   TryParseTimeSpan(str, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="string"/>.
        /// </summary>
        /// <param name="propertyExpression">Expected property</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetString(Expression<Func<TRowModel, object>> propertyExpression, out string value)
        {
            var propertyInfo = ExpressionExtensions.GetPropertyInfo(propertyExpression);

            if (!ExpectedColumns.Values.Contains(propertyInfo))
            {
                throw new InvalidOperationException($"Property '{propertyInfo.Name}' is not an expected column in the spreadsheet.");
            }

            var columnHeader = ExpectedColumns.Single(kvp => kvp.Value == propertyInfo).Key;

            return TryGetString(columnHeader, out value);
        }

        /// <summary>
        /// Try get cell value from the current row as a <see cref="string"/>.
        /// </summary>
        /// <param name="columnName">Column Header</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetString(string columnName, out string value)
        {
            var foundCell = TryGetCell(columnName, out var cell);
            value = cell?.Value?.ToString() ?? string.Empty;
            return foundCell;
        }

        /// <summary>
        /// Attempts to parse a <see cref="DateTime"/> from a <see cref="string"/> using a predefined set of datetime formats.
        /// </summary>
        /// <param name="valueString">Formatted string</param>
        /// <param name="value">Output value</param>
        /// <returns><see cref="true"/> if successful</returns>
        protected bool TryParseDateTime(string valueString, out DateTime value)
        {
            // Could be datetime, or number format
            if (DateTime.TryParseExact(valueString, DateTimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out value))
            {
                return true;
            }
            else if (valueString.IsNumeric() &&
                double.TryParse(valueString, out var oaDate) &&
                oaDate > 0)
            {
                value = DateTimes.FromOADatePrecise(oaDate);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Attempts to parse a <see cref="DateTimeOffset"/> from a <see cref="string"/> using a predefined set of datetime formats.
        /// </summary>
        /// <param name="valueString">Formatted string</param>
        /// <param name="value">Output value</param>
        /// <returns><see cref="true"/> if successful</returns>
        protected bool TryParseDateTimeOffset(string valueString, out DateTimeOffset value)
        {
            // Could be datetime, or number format
            if (DateTimeOffset.TryParseExact(valueString, DateTimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out value))
            {
                return true;
            }
            else if (valueString.IsNumeric() &&
                double.TryParse(valueString, out var oaDate) &&
                oaDate > 0)
            {
                value = new DateTimeOffset(DateTimes.FromOADatePrecise(oaDate));
                return true;
            }

            return false;
        }

        /// <summary>
        /// Attempts to parse a <see cref="TimeSpan"/> from a <see cref="string"/>.
        /// </summary>
        /// <param name="valueString">Formatted string</param>
        /// <param name="value">Output value</param>
        /// <returns><see cref="true"/> if successful</returns>
        protected bool TryParseTimeSpan(string valueString, out TimeSpan value)
        {
            // Could be datetime, timespan, or number format
            if (TimeSpan.TryParseExact(valueString, "g", CultureInfo.InvariantCulture, out value))
            {
                return true;
            }
            else if (TryParseDateTime(valueString, out var dt))
            {
                value = dt.TimeOfDay;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Override to add further validation logic for the current row.
        /// This function will not be called if the row is already considered to be invalid after parsing.
        /// </summary>
        /// <param name="row">A parsed row model</param>
        /// <returns>An enumerable of identified validation errors</returns>
        protected virtual IEnumerable<IImportError> ValidateRow(TRowModel row)
        {
            return Enumerable.Empty<IImportError>();
        }

        /// <summary>
        /// Attempts to parse the property value to be from the cell.
        /// This is initally done by TableRow.FromOpenXml(), however that can't support enum conversion yet.
        /// </summary>
        /// <param name="columnName">Column name for field errors</param>
        /// <param name="propertyType">Property type to convert value to</param>
        /// <param name="cell">Spreadsheet cell containing value</param>
        /// <param name="importErrors">Row import errors</param>
        /// <returns>Parsed cell value</returns>
        private object ParseValue(string columnName, Type propertyType, bool isNullable, TableCell cell, List<IImportError> importErrors)
        {
            // Don't bother parsing if null
            if (cell?.Value == null)
            {
                return null;
            }

            // If it is already the correct type or previously parsed then use it by TableRow
            if (cell.Value.GetType() == propertyType)
            {
                return cell.Value;
            }

            // If a DateTimeOffset is expected and we have a DateTime then wrap the DateTime and return it straight away
            if (cell.Value.GetType() == typeof(DateTime) &&
                propertyType == typeof(DateTimeOffset))
            {
                return new DateTimeOffset((DateTime)cell.Value);
            }

            var valueString = cell.GetValue();

            // Skip parsing if empty value and nullable type
            if (isNullable && string.IsNullOrWhiteSpace(valueString))
            {
                return null;
            }

            // Fallback parser based on the output property type for when the number format isn't parsed
            if (propertyType == typeof(DateTime))
            {
                if (TryParseDateTime(valueString, out var dt))
                {
                    return dt;
                }
            }
            else if (propertyType == typeof(DateTimeOffset))
            {
                if (TryParseDateTimeOffset(valueString, out var dto))
                {
                    return dto;
                }
            }
            else if (propertyType == typeof(TimeSpan))
            {
                if (TryParseTimeSpan(valueString, out var t))
                {
                    return t;
                }
            }
            else if (propertyType == typeof(decimal))
            {
                if (decimal.TryParse(valueString, out var dec))
                {
                    return dec;
                }
            }
            else if (propertyType == typeof(double))
            {
                if (double.TryParse(valueString, out var doub))
                {
                    return doub;
                }
            }
            else if (propertyType == typeof(float))
            {
                if (float.TryParse(valueString, out var f))
                {
                    return f;
                }
            }
            else if (propertyType == typeof(int))
            {
                if (int.TryParse(valueString, out var i))
                {
                    return i;
                }
            }
            else if (propertyType == typeof(bool))
            {
                if (Bool.TryParse(valueString, out var b))
                {
                    return b;
                }
            }
            else if (propertyType.IsEnum)
            {
                if (EnumMember.TryParse(propertyType, valueString, out var enumValue))
                {
                    return enumValue;
                }

                // Override the default import error to include possible values
                importErrors.Add(new FieldParseError(GetRowNumber(), columnName, valueString, EnumMember.Options(propertyType).ToArray()));
                return null;
            }

            importErrors.Add(new FieldParseError(GetRowNumber(), columnName, valueString));

            return null;
        }
    }
}
