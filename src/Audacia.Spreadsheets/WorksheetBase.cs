using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace Audacia.Spreadsheets
{
    public abstract class WorksheetBase
    {
        public string SheetName { get; set; }
        public FreezePane FreezePane { get; set; }
        public SheetStateValues Visibility { get; set; } = SheetStateValues.Visible;
        public bool ShowGridLines { get; set; } = false;
        public bool HasAutofilter { get; set; } = false;
        public WorksheetProtection WorksheetProtection { get; set; }
        public List<StaticDropdown> StaticDataValidations { get; } = new List<StaticDropdown>();
        public List<DependentDropdown> DependentDataValidations { get; } = new List<DependentDropdown>();

        /// <summary>
        /// Sets Visibility to Hidden.
        /// </summary>
        /// <param name="completelyHidden">If true the worksheet will not be visible from Excel</param>
        public void Hide(bool completelyHidden = false)
        {
            Visibility = completelyHidden
                ? SheetStateValues.VeryHidden
                : SheetStateValues.Hidden;
        }

        protected abstract void WriteSheetData(SharedDataTable sharedData, OpenXmlWriter writer);

        public void Write(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            // Create a worksheet
            writer.WriteStartElement(new OpenXmlWorksheet());

            // Write meta data for the worksheet
            // Sheet view, Columns, and Sheet Data should only ever be written once per worksheet
            AddSheetView(writer);
            
            var allTables = this.GetTables().ToArray();

            DefineColumnsIfRequired(allTables);

            AddColumns(allTables, writer);

            // Create a place to store sheet data
            writer.WriteStartElement(new SheetData());

            // write the sheet data
            WriteSheetData(sharedData, writer);

            // Close SheetData tag
            writer.WriteEndElement();
            
            AddProtection(writer);
            
            // We don't currently support autofilters for multi-table worksheets
            if (HasAutofilter && allTables.Any())
            {
                AddAutoFilter(allTables.First(), sharedData.DefinedNames, writer);
            }

            // Add data validation if required
            var dataValidations = new DataValidations();
            
            // Add Static Data Validation
            if (StaticDataValidations != null && StaticDataValidations.Any())
            {
                foreach (var val in StaticDataValidations)
                {
                    val.Write(dataValidations);
                }
            }
            
            // Add Dynamic Data Validation
            if (DependentDataValidations != null && DependentDataValidations.Any())
            {
                foreach (var val in DependentDataValidations)
                {
                    val.Write(dataValidations);
                }
            }

            //  Only add validation if dataValidations has Descendants
            if (dataValidations.Descendants<DataValidation>().Any())
            {
                writer.WriteElement(dataValidations);
            }
            
            // Close the worksheet
            writer.WriteEndElement();
        }

        protected void AddAutoFilter(Table table, DefinedNames definedNames, OpenXmlWriter writer)
        {
            if (table.IncludeHeaders && table.Rows.Any())
            {
                var firstCell = new CellReference(table.StartingCellRef);

                // Step over the subtotal row, onto the header row
                if (table.Columns.Any(h => h.DisplaySubtotal))
                {
                    firstCell.NextRow();
                }

                var lastCell = firstCell.MutateBy(table.Columns.Count - 1, table.Rows.Count());

                // Selects All Column Headers & Data
                var cellReference = $"{firstCell}:{lastCell}";

                var filter = new AutoFilter { Reference = cellReference };

                // Excel 2013 Requires a Defined Name to be able to sort using the AutoFilter
                var dn = new DefinedName
                {
                    Text = $"'{SheetName}'!${cellReference}",
                    Name = "_xlnm._FilterDatabase", // Don't rename this or else Excel 2013 will crash
                    LocalSheetId = (uint) 0,
                    Hidden = true
                };
                
                definedNames.Append(dn);
                writer.WriteElement(filter);
            }
        }

        private void DefineColumnsIfRequired(IEnumerable<Table> tables)
        {
            // If the developer has not added any column headers, then we need to do this.
            // We use TableColumns to define column metadata in OpenXML and also to write cell content.
            // They do not get to benefit from number formats because they haven't defined their columns themselves.
            foreach (var table in tables)
            {
                if (!table.IncludeHeaders && !table.Columns.Any())
                {
                    var maxCells = table.Rows.Max(r => r.Cells.Count);
                    table.Columns = Enumerable.Range(0, maxCells)
                        .Select(_ => new TableColumn())
                        .ToList();
                }
            }
        }

        /// <summary>
        /// Defines column metadata in the spreadsheet.
        /// This is specifically for OpenXML columns & defining column widths.
        /// </summary>
        /// <param name="tables"> A List of tables to add the columns too.</param>
        /// <param name="writer"> A XML writer for writing worksheets.</param>
        protected static void AddColumns(IList<Table> tables, OpenXmlWriter writer)
        {
            var columns = new Columns();
            writer.WriteStartElement(columns);

            // Find the table with the most columns and get the total columns
            var maxColumnCount = tables.Max(table => table.Columns.Count);

            const double maxDigitWidth = 11D;

            for (var columnIndex = 0; columnIndex < maxColumnCount; columnIndex++)
            {
                // Find the max cell width from all tables with the column
                var tableMaxCellWidth = tables
                    .Where(table => columnIndex < table.Columns.Count)
                    .Select(table => new {Table = table, MaxCellWidth = table.GetMaxCharacterWidth(columnIndex)})
                    .OrderByDescending(tableMaxCellWidth => tableMaxCellWidth.MaxCellWidth)
                    .FirstOrDefault();

                // width = Truncate([{Number of Characters} * {Maximum Digit Width} + {20 pixel padding}]/{Maximum Digit Width}*256)/256
                const int cellPadding = 20;
                const int widthConstant = 256;
                var cellWidth = tableMaxCellWidth!.MaxCellWidth * maxDigitWidth;
                var totalCellWidth = cellWidth + cellPadding;
                var width = Math.Truncate(totalCellWidth / maxDigitWidth * widthConstant) / widthConstant;

                // Limit the column width to 75...
                if (width > 75)
                {
                    width = 75;
                }

                // To adjust for font size.
                var factor = (tableMaxCellWidth.Table.HeaderStyle?.FontSize ?? 11) / maxDigitWidth;

                var colWidth = new DoubleValue(width * factor);
                var column = new Column()
                {
                    Min = Convert.ToUInt32(columnIndex + 1),
                    Max = Convert.ToUInt32(columnIndex + 1),
                    CustomWidth = true,
                    BestFit = true,
                    Width = colWidth,
                };
                writer.WriteElement(column);
            }

            writer.WriteEndElement();
        }

        private void AddProtection(OpenXmlWriter writer)
        {
#pragma warning disable SA1101
            if (WorksheetProtection == null)
#pragma warning restore SA1101
            {
                return;
            }

            var sheetProtection = new SheetProtection
            {
                Objects = true,
                Scenarios = true,
                Sheet = true,
                InsertColumns = !WorksheetProtection.CanAddOrDeleteColumns,
                DeleteColumns = !WorksheetProtection.CanAddOrDeleteColumns,
                InsertRows = !WorksheetProtection.CanAddOrDeleteRows,
                DeleteRows = !WorksheetProtection.CanAddOrDeleteRows,
                //  A value of 0 allows AutoFiltering
                AutoFilter = !WorksheetProtection.AllowAutoFilter,
                //  A value of 0 allows Sorting
                Sort = !WorksheetProtection.AllowSort,
                FormatRows = !WorksheetProtection.AllowFormatRows,
                FormatColumns = !WorksheetProtection.AllowFormatColumns,
                FormatCells = !WorksheetProtection.AllowFormatCells
            };

            if (!string.IsNullOrWhiteSpace(WorksheetProtection.Password))
            {
                // NOTE: We cannot use Workbook protection, as the resulting OpenXML file is marked as corrupted
                // by OpenXML when attempting to open it - the Productivity tool does the same thing.
                // So we'll just do worksheet protection
                sheetProtection.Password = HexPasswordConversion(WorksheetProtection.Password);
                sheetProtection.SelectLockedCells = false;
            }
        }

        protected void AddSheetView(OpenXmlWriter writer)
        {
            writer.WriteStartElement(new SheetViews());
            var sheetView = new SheetView
            {
                ShowGridLines = ShowGridLines,
                WorkbookViewId = 0U
            };
            
            FreezePane?.Write(sheetView);

            writer.WriteElement(sheetView);
            writer.WriteEndElement();
        }

        protected static int GetMaxRowWidth(WorksheetPart worksheetPart)
        {
            var maxWidth = 0;
            var rows = worksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().ToList();

            for (var i = 1; i < rows.Count; i++)
            {
                var row = rows[i];
                var lastCell = row.Elements<Cell>().LastOrDefault();
                if (lastCell == default(Cell)) continue;

                var rowIndex = lastCell.CellReference.Value.GetRowNumber();
                if (rowIndex > maxWidth)
                {
                    maxWidth = (int)rowIndex;
                }
            }

            return maxWidth;
        }
        
        private static HexBinaryValue HexPasswordConversion(string password)
        {
            if (string.IsNullOrWhiteSpace(password))
            {
                throw new ArgumentException("Cannot convert an empty password");
            }

            var passwordCharacters = System.Text.Encoding.ASCII.GetBytes(password);
            var hash = 0;
            if (passwordCharacters.Length > 0)
            {
                var charIndex = passwordCharacters.Length;

                while (charIndex-- > 0)
                {
                    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                    hash ^= passwordCharacters[charIndex];
                }
                // Main difference from spec, also hash with char count
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters.Length;
                hash ^= (0x8000 | ('N' << 8) | 'K');
            }

            return Convert.ToString(hash, 16).ToUpperInvariant();
        }
    }
}
