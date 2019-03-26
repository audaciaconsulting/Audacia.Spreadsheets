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
    public class Worksheet
    {
        public string SheetName { get; set; }
        public Table Table { get; set; }
        public FreezePane FreezePane { get; set; }
        public SheetStateValues Visibility { get; set; } = SheetStateValues.Visible;
        public bool ShowGridLines { get; set; } = false;
        
        public bool ShowBorders { get; set; } = true;
        public WorksheetProtection WorksheetProtection { get; set; }
        public List<StaticDropdown> StaticDataValidations { get; } = new List<StaticDropdown>();
        public List<DependentDropdown> DependentDataValidations { get; } = new List<DependentDropdown>();

        public void Write(WorksheetPart worksheetPart, SharedDataTable sharedData)
        {
            var writer = OpenXmlWriter.Create(worksheetPart);

            writer.WriteStartElement(new OpenXmlWorksheet());

            AddSheetView(writer);
            AddColumns(Table, writer);

            writer.WriteStartElement(new SheetData());

            Table.Write(sharedData, writer);

            // Auto Filter for a single table on the worksheet
            AddAutoFilter(Table, sharedData.DefinedNames, writer);

            writer.WriteEndElement(); // Sheet Data

            DataValidations dataValidations = new DataValidations();
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


            writer.WriteEndElement(); // Worksheet

            writer.Close();
            
            AddProtection(worksheetPart);
        }

        private void AddAutoFilter(Table table, DefinedNames definedNames, OpenXmlWriter writer)
        {
            if (table.IncludeHeaders && table.Rows.Any())
            {
                var firstCell = new CellReference(table.StartingCellRef);

                // Step over the subtotal row, onto the header row
                if (table.Columns.Any(h => h.DisplaySubtotal))
                {
                    firstCell.NextRow();
                }

                var lastCell = firstCell.MutateBy(table.Columns.Count - 1, table.Rows.Count);

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
        
        private void AddColumns(Table table, OpenXmlWriter writer)
        {
            writer.WriteStartElement(new Columns());
            const double maxWidth = 11D;

            for (var i = 0; i < table.Columns.Count; i++)
            {
                var item = Table.GetMaxCharacterWidth(table, i);

                //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {20 pixel padding}]/{Maximum Digit Width}*256)/256
                var width = Math.Truncate((item * maxWidth + 20) / maxWidth * 256) / 256;

                //  To adjust for font size.
                var factor = (table.HeaderStyle?.FontSize ?? 11) / maxWidth;

                var colWidth = (DoubleValue)(width * factor);

                writer.WriteElement(new Column
                {
                    Min = Convert.ToUInt32(i + 1),
                    Max = Convert.ToUInt32(i + 1),
                    CustomWidth = true,
                    BestFit = true,
                    Width = colWidth
                });
            }

            writer.WriteEndElement();
        }
        
        private void AddProtection(WorksheetPart worksheetPart)
        {
            if (WorksheetProtection == null)
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
                DeleteRows = !WorksheetProtection.CanAddOrDeleteRows
            };

            if (!string.IsNullOrWhiteSpace(WorksheetProtection.Password))
            {
                // NOTE: We cannot use Workbook protection, as the resulting OpenXML file is marked as corrupted
                // by OpenXML when attempting to open it - the Productivity tool does the same thing.
                // So we'll just do worksheet protection
                sheetProtection.Password = HexPasswordConversion(WorksheetProtection.Password);
            }

            var pRanges = new ProtectedRanges();

            foreach (var protectedRange in WorksheetProtection.EditableCellRanges)
            {
                var pRange = new ProtectedRange();
                var lValue = new ListValue<StringValue> { InnerText = protectedRange };

                pRange.SequenceOfReferences = lValue;
                pRange.Name = "not allow editing";
                pRanges.Append(pRange);
            }

            //These are the cells that are editable
            var pageM = worksheetPart.Worksheet.GetFirstChild<PageMargins>();
            worksheetPart.Worksheet.InsertBefore(sheetProtection, pageM);
            worksheetPart.Worksheet.InsertBefore(pRanges, pageM);
        }

        private void AddSheetView(OpenXmlWriter writer)
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

        public static Worksheet FromOpenXml(Sheet worksheet, SpreadsheetDocument spreadSheet, bool includeHeaders, bool hasSubtotals)
        {
            var worksheetPart = (WorksheetPart) spreadSheet.WorkbookPart.GetPartById(worksheet.Id);

            var table = new Table
            {
                StartingCellRef = "A1",
                IncludeHeaders = includeHeaders,
                HeaderStyle = null
            };

            var startingRowIndex = 0;
            if (hasSubtotals && includeHeaders)
            {
                // Rows start at i = 2 if this sheet was also exported with subtotals 
                startingRowIndex += 2;
            } 
            else if (includeHeaders)
            {
                // Rows start at i = 1 to skip header row IF headers are included
                startingRowIndex += 1;
            }

            if (includeHeaders)
            {
                var columns = TableColumn.FromOpenXml(worksheetPart, spreadSheet, hasSubtotals);
                table.Columns.AddRange(columns);
            }

            var maxRowWidth = includeHeaders ? table.Columns.Count : GetMaxRowWidth(worksheetPart);

            var rows = TableRow.FromOpenXml(worksheetPart, spreadSheet, maxRowWidth, startingRowIndex);
            table.Rows.AddRange(rows);

            return new Worksheet
            {
                SheetName = worksheet.Name,
                Table = table
            };
        }
        
        private static int GetMaxRowWidth(WorksheetPart worksheetPart)
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
