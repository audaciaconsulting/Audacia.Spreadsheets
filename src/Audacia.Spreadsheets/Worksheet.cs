using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace Audacia.Spreadsheets
{
    public class Worksheet : WorksheetBase
    {
        public Table Table { get; set; }
        
        protected override void WriteSheetContent(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            AddSheetView(writer);
            AddColumns(Table, writer);
            
            Table.Write(sharedData, writer);
            
            if (HasAutofilter)
            {
                AddAutoFilter(Table, sharedData.DefinedNames, writer);
            }
            
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

            // Force enumeration of the content when reading the worksheet, otherwise the spreadsheet is disposed before we can read the data.
            table.Rows = TableRow.FromOpenXml(worksheetPart, spreadSheet, maxRowWidth, startingRowIndex).ToArray();

            return new Worksheet
            {
                SheetName = worksheet.Name,
                Table = table,
                Visibility = worksheet.State ?? SheetStateValues.Visible
            };
        }
    }
}
