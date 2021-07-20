using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
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
            "dd/MM/yyyy HH:mm:ss",
            "dd/MM/yyyy",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd"
        };

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
        /// Current worksheet being parsed by the importer.
        /// </summary>
        protected Worksheet Worksheet { get; private set; }

        /// <summary>
        /// Set to true before manually mapping columns to skip automatic column header mapping.
        /// To be used in situations where a spreadsheet has no column headers.
        /// </summary>
        public bool OverrideSpreadsheetColumnMapping { get; set; }

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
            if (OverrideSpreadsheetColumnMapping)
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
        /// If there are no column headers on the sheet, you will need to manuall configure ColumnMap to use this.
        /// </summary>
        /// <param name="worksheet">Worksheet to be parsed</param>
        /// <param name="ignoreProperties">Properties to ignore when generating expected column headers</param>
        public virtual IEnumerable<ImportRow<TRowModel>> ParseWorksheet(WorksheetBase worksheet, params string[] ignoreProperties)
        {
            // We only support single worksheets
            Worksheet = worksheet as Worksheet;
            if (Worksheet == null)
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
                // Check for duplicate column names in spreadsheet
                var duplicateColumnNames = Worksheet.Table.Columns
                    .Where(c => !string.IsNullOrWhiteSpace(c.Name))
                    .Where(c => Worksheet.Table.Columns.Count(tc => tc.Name == c.Name) > 1)
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
                SpreadsheetColumns = Worksheet.Table.Columns.ToDictionary();

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
            foreach (var row in Worksheet.Table.Rows)
            {
                CurrentRow = row;
                var rowParseErrors = ParseRow(out var rowModel);

                // Allow for custom row validation if inherited
                var customValidationErrors = rowParseErrors.Any()
                    ? Enumerable.Empty<IImportError>()
                    : ValidateRow(rowModel);

                var importModel = new ImportRow<TRowModel>
                {
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
        /// Handles the parsing of the CurrentRow, should add validation errors where necessary.
        /// </summary>
        protected virtual IList<IImportError> ParseRow(out TRowModel model)
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
                    rowErrors.Add(new FieldMissingError(CurrentRow.Id.Value, columnName));
                    continue;
                }

                var propertyType = expectedProperty.Value.PropertyType;
                var underlyingType = propertyType.GetUnderlyingTypeIfNullable();
                var isNullable = expectedProperty.Value.PropertyType.IsNullable();

                // Get the cell value and parse it
                var parsedValue = ParseValue(columnName, underlyingType, isNullable, cell, rowErrors);

                // Set cell value
                if (isNullable || parsedValue != null)
                {
                    expectedProperty.Value.SetValue(model, parsedValue);
                }
                
            }

            return rowErrors;
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

            var rowId = CurrentRow.Id.Value;
            var valueString = cell.GetValue();

            // Skip parsing if empty value and nullable type
            if (isNullable && string.IsNullOrWhiteSpace(valueString))
            {
                return null;
            }

            // Fallback parser based on the output property type for when the number format isn't parsed
            if (propertyType == typeof(DateTime))
            {
                // Could be datetime, or number format
                if (DateTime.TryParseExact(valueString, DateTimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                {
                    return dt;
                }
                else
                {
                    try
                    {
                        return DateTimes.FromOADatePrecise(double.Parse(valueString));
                    }
                    catch
                    {
                        // Import error is raised below
                    }
                }
            }
            else if (propertyType == typeof(DateTimeOffset))
            {
                // Could be datetime, or number format
                if (DateTimeOffset.TryParseExact(valueString, DateTimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dtoff))
                {
                    return dtoff;
                }
                else 
                {
                    try
                    {
                        var datetime = DateTimes.FromOADatePrecise(double.Parse(valueString));
                        return new DateTimeOffset(datetime);
                    }
                    catch
                    {
                        // Import error is raised below
                    }
                }
            }
            else if (propertyType == typeof(TimeSpan))
            {
                // Could be datetime, timespan, or number format
                if (TimeSpan.TryParseExact(valueString, "g", CultureInfo.InvariantCulture, out var t))
                {
                    return t;
                }
                else if (DateTime.TryParseExact(valueString, DateTimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                {
                    return new TimeSpan(dt.Hour, dt.Minute, dt.Second, dt.Millisecond);
                }
                else
                {
                    try
                    {
                        var datetime = DateTimes.FromOADatePrecise(double.Parse(valueString));
                        return datetime.TimeOfDay;
                    }
                    catch
                    {
                        // Import error is raised below
                    }
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
                try
                {
                    return Enum.Parse(propertyType, valueString, ignoreCase: true);
                }
                catch
                {
                    // No option to tryparse without being strongly typed
                    // Import error is raised below
                }
            }

            importErrors.Add(new FieldParseError(rowId, columnName, valueString));

            return null;
        }

        /// <summary>
        /// Try get cell data from the current row.
        /// Cell.Value can be a string, decimal, or DateTime.
        /// Cell.FillColour is also parsed.
        /// </summary>
        /// <param name="columnName">Column Header</param>
        /// <param name="cell">Cell Data</param>
        protected bool TryGetCell(string columnName, out TableCell cell)
        {
            cell = null;

            if (!SpreadsheetColumns.ContainsKey(columnName)) return false;

            var columnIndex = SpreadsheetColumns[columnName];

            if (CurrentRow.Cells.Count <= columnIndex) return false;

            cell = CurrentRow.Cells[columnIndex];

            return true;
        }

        /// <summary>
        /// Try get cell value from the current row as a string.
        /// </summary>
        /// <param name="columnName">Column Header</param>
        /// <param name="value">Cell Value</param>
        protected bool TryGetValueStr(string columnName, out string value)
        {
            var foundCell = TryGetCell(columnName, out var cell);
            value = cell?.Value.ToString() ?? string.Empty;
            return foundCell;
        }

        /// <summary>
        /// Gets the cell value by name from the provided fields.
        /// </summary>
        /// <param name="fields">Cell values</param>
        /// <param name="columnName">Column header</param>
        protected string GetValue(string[] fields, string columnName)
        {
            if (!SpreadsheetColumns.ContainsKey(columnName)) return string.Empty;

            var columnIndex = SpreadsheetColumns[columnName];

            if (fields.Length <= columnIndex) return string.Empty;

            return fields[columnIndex];
        }
    }
}
