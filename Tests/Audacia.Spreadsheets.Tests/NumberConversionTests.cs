using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Audacia.Spreadsheets.Tests
{
    /// <summary>
    /// Unit tests covering number conversion logic.
    /// </summary>
    public class NumberConversionTests
    {
        /// <summary>
        /// Asserts that given a decimal value, the conversion logic correctly handles floating point errors. 
        /// </summary>
        [Fact]
        public void FloatingPointErrorsAreHandledDuringWorksheetParsing()
        {
            const decimal value = 1.1m;

            // Arrange: generate a spreadsheet with a single worksheet, single column, and a single row with a value of 1.1.
            using var stream = GenerateDecimalValueSpreadsheetStream(value);

            // Act: read the value using TableRow.FromOpenXml
            using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
            var workbookPart = spreadsheetDocument.WorkbookPart!.WorksheetParts.First();
            var rows = TableRow.FromOpenXml(workbookPart, spreadsheetDocument, 1).ToList();
            var cellValue = rows.Last().Cells.First().Value!.ToString();

            Assert.True(decimal.TryParse(cellValue, out var result));
            Assert.Equal(value, result);
        }

        /// <summary>
        /// Generates a sample spreadsheet with a single decimal field, and returns it as a <see cref="Stream"/>.
        /// </summary>
        private static MemoryStream GenerateDecimalValueSpreadsheetStream(decimal cellValue)
        {
            var stream = new MemoryStream();

            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                var sharedStringTable = new SharedStringTable();
                sharedStringTable.Append(new SharedStringItem(new Text("Example Column")));
                sharedStringTable.Append(new SharedStringItem(new Text(cellValue.ToString())));
                sharedStringPart.SharedStringTable = sharedStringTable;
                sharedStringPart.SharedStringTable.Save();

                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet(
                    new Fonts(new Font()),
                    new Fills(new Fill()),
                    new Borders(new Border()),
                    new CellFormats()
                );

                stylesPart.Stylesheet.Save();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();

                var headerRow = new Row() { RowIndex = 1 };
                var cell = new Cell()
                {
                    CellReference = "A1",
                    DataType = CellValues.SharedString,
                    CellValue = new CellValue("0")
                };

                headerRow.Append(cell);
                sheetData.Append(headerRow);

                var dataRow = new Row() { RowIndex = 2 };
                dataRow.Append(
                    new Cell
                    {
                        CellReference = "A2",
                        DataType = CellValues.Number,
                        CellValue = new CellValue(cellValue)
                    });

                sheetData.Append(dataRow);

                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);
                worksheetPart.Worksheet.Save();

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sample Worksheet"
                });

                workbookPart.Workbook.Save();
            }

            stream.Position = 0;

            return stream;
        }
    }
}
