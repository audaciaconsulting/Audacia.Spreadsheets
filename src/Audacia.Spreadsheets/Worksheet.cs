using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Core.Extensions;
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
        public IEnumerable<Table> Tables { get; set; }
        public WorksheetProtection WorksheetProtection { get; set; }

        public void Write(WorksheetPart worksheetPart, SharedData sharedData)
        {
            var writer = OpenXmlWriter.Create(worksheetPart);

            foreach (var table in Tables)
            {
                writer.WriteStartElement(new OpenXmlWorksheet());

                SpreadsheetBuilderHelper.AddSheetView(writer, table.FreezeTopRows);
                SpreadsheetBuilderHelper.AddColumns(writer, table);

                writer.WriteStartElement(new SheetData());

                SpreadsheetBuilderHelper.Insert(table,
                    sharedData.Stylesheet,
                    sharedData.CellFormats,
                    sharedData.FillColours,
                    sharedData.TextColours,
                    sharedData.Fonts,
                    worksheetPart,
                    writer);

                writer.WriteEndElement(); // Sheet Data

                // Auto Filter for a single table on the worksheet
                AddAutoFilter(writer, table, sharedData.DefinedNames);

                writer.WriteEndElement(); // Worksheet
            }

            writer.Close();

            AddProtection(worksheetPart);
        }

        private void AddAutoFilter(OpenXmlWriter writer, Table table, DefinedNames definedNames)
        {
            // TODO JP: make this more efficient, only add filters to the first table that requests it
            if (table.IncludeHeaders && table.Rows.Any())
            {
                // 'A1'
                var initialCellRef = !string.IsNullOrWhiteSpace(table.StartingCellRef)
                    ? table.StartingCellRef
                    : "A1";

                // 'A'
                var firstColumnRef = initialCellRef.GetReferenceColumnIndex();

                // '1' or '2' - Handles Rollups above Cell Headers
                var firstRowRef = initialCellRef.GetReferenceRowIndex() +
                                  (table.Columns.Any(h => h.ColumnRollup) ? 1 : 0);

                var lastColumnRef =
                    (firstColumnRef.GetColumnNumber() +
                     table.Columns.Count - 1)
                    .GetExcelColumnName();

                var lastRowRef = firstRowRef + table.Rows.Count;

                // Selects All Column Headers & Data
                var cellReference = $"{firstColumnRef}{firstRowRef}:{lastColumnRef}{lastRowRef}";

                var filter = new AutoFilter {Reference = cellReference};

                // Excel 2013 Requires a Defined Name to be able to sort using the AutoFilter
                var dn = new DefinedName
                {
                    Text = $"'{SheetName}'!${firstColumnRef}${firstRowRef}:${lastColumnRef}${lastRowRef}",
                    Name = "_xlnm._FilterDatabase", // Don't rename this or else Excel 2013 will crash
                    LocalSheetId = (uint) 0,
                    Hidden = true
                };
                
                definedNames.Append(dn);
                writer.WriteElement(filter);
            }
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
                DeleteRows = !WorksheetProtection.CanAddOrDeleteRows,
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
