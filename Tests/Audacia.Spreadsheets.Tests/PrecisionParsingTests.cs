using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Audacia.Spreadsheets.Tests
{
    public class PrecisionParsingTests
    {
        [Fact]
        public void Excess_precision_numbers_are_normalized()
        {
            byte[] bytes;
            using (var ms = new MemoryStream())
            {
                using (var spreadsheet = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
                {
                    var workbookPart = spreadsheet.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                    sheets.Append(sheet);

                    var row = new Row();
                    var cell = new Cell { CellReference = "A1", DataType = CellValues.Number };
                    cell.CellValue = new CellValue("1.10000000001");
                    row.Append(cell);
                    sheetData.Append(row);
                }

                bytes = ms.ToArray();
            }

            var worksheet = Spreadsheet.FromBytes(bytes).Worksheets[0];
            var value = (decimal)worksheet.Table.Rows.First().Cells.First().Value!;
            Assert.Equal(1.1m, value);
        }
    }
}

