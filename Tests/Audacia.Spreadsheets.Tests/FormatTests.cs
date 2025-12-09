using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using Xunit;

namespace Audacia.Spreadsheets.Tests
{
    /// <summary>
    /// Unit tests covering number conversion logic.
    /// </summary>
    public class FormatTests
    {
        /// <summary>
        /// Asserts that MergeCells are added to the OpenXML document when set in the Worksheet model. 
        /// </summary>
        [Fact]
        public void MergeCellsAreAdded()
        {
            const string cellRef = "A1:B1";
            var spreadsheet = GetSpreadsheet();

            spreadsheet.Worksheets.First().MergeCells = new List<string>() { cellRef };

            using var stream = new MemoryStream();
            spreadsheet.Write(stream);

            using var spreadSheet = SpreadsheetDocument.Open(stream, false);
            var workbookPart = spreadSheet.WorkbookPart;
            var worksheetPart = workbookPart!.WorksheetParts.First();
            var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            Assert.NotNull(mergeCells);

            var mergeCell = mergeCells?.Elements<MergeCell>().FirstOrDefault();
            Assert.NotNull(mergeCell);
            Assert.Equal(cellRef, mergeCell?.Reference?.Value);
        }

        /// <summary>
        /// Asserts that MergeCells are not added to the OpenXML document when not set in the Worksheet model. 
        /// </summary>
        [Fact]
        public void MergeCellsAreNotAdded()
        {
            var spreadsheet = GetSpreadsheet();

            using var stream = new MemoryStream();
            spreadsheet.Write(stream);

            using var spreadSheet = SpreadsheetDocument.Open(stream, false);
            var workbookPart = spreadSheet.WorkbookPart;
            var worksheetPart = workbookPart!.WorksheetParts.First();
            var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            Assert.Null(mergeCells);
        }

        /// <summary>
        /// Asserts that text is horizontally centred in the OpenXML document when set in the TableCell model. 
        /// </summary>
        [Fact]
        public void TextIsHorizontallyCentredInCell()
        {
            const HorizontalAlignmentValues alignment = HorizontalAlignmentValues.Center;
            var spreadsheet = GetSpreadsheet();

            spreadsheet.Worksheets.First().GetTable().Rows.First().Cells.First().AlignHorizontal = alignment;

            using var stream = new MemoryStream();
            spreadsheet.Write(stream);

            using var spreadSheet = SpreadsheetDocument.Open(stream, false);
            var workbookPart = spreadSheet.WorkbookPart;
            var worksheetPart = workbookPart!.WorksheetParts.First();
            var stylesheet = spreadSheet.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
            var cellFormats = stylesheet?.CellFormats;

            var row = worksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().First();
            var cell = row.Elements<Cell>().First();
            var cellStyleIndex = cell.StyleIndex?.Value;
            var cellFormat = cellFormats?.ElementAt((int)cellStyleIndex!.Value) as DocumentFormat.OpenXml.Spreadsheet.CellFormat;

            Assert.NotNull(cellFormat);
            Assert.NotNull(cellFormat?.Alignment?.Horizontal?.Value);
            Assert.Equal(cellFormat?.Alignment?.Horizontal?.Value, alignment);
        }

        /// <summary>
        /// Asserts that text is aligned left in the OpenXML document by default. 
        /// </summary>
        [Fact]
        public void TextIsHorizontallyLeftAlignedInCellByDefault()
        {
            var spreadsheet = GetSpreadsheet();

            using var stream = new MemoryStream();
            spreadsheet.Write(stream);

            using var spreadSheet = SpreadsheetDocument.Open(stream, false);
            var workbookPart = spreadSheet.WorkbookPart;
            var worksheetPart = workbookPart!.WorksheetParts.First();
            var stylesheet = spreadSheet.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
            var cellFormats = stylesheet?.CellFormats;

            var row = worksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().First();
            var cell = row.Elements<Cell>().First();
            var cellStyleIndex = cell.StyleIndex?.Value;
            var cellFormat = cellFormats?.ElementAt((int)cellStyleIndex!.Value) as DocumentFormat.OpenXml.Spreadsheet.CellFormat;

            Assert.NotNull(cellFormat);
            Assert.NotNull(cellFormat?.Alignment?.Horizontal?.Value);
            Assert.Equal(cellFormat?.Alignment?.Horizontal?.Value, HorizontalAlignmentValues.Left);
        }

        /// <summary>
        /// Asserts that a column is hidden in the OpenXML document when IsHidden is true in the TableColumn model. 
        /// </summary>
        [Fact]
        public void ColumnIsHiddenWhenSetInModel()
        {
            var spreadsheet = GetSpreadsheet();

            spreadsheet.Worksheets.First().GetTable().Columns.First().IsHidden = true;

            using var stream = new MemoryStream();
            spreadsheet.Write(stream);

            using var spreadSheet = SpreadsheetDocument.Open(stream, false);
            var workbookPart = spreadSheet.WorkbookPart;
            var worksheetPart = workbookPart!.WorksheetParts.First();

            var columns = worksheetPart.Worksheet.Elements<Columns>().First();
            var hiddenColumn = columns.Elements<Column>().First();
            
            Assert.NotNull(hiddenColumn);
            Assert.True(hiddenColumn.Hidden?.Value);

            var visibleColumn = columns.Elements<Column>().ToArray()[1];
            Assert.NotNull(visibleColumn);
            Assert.False(visibleColumn.Hidden?.Value);
        }

        private static Spreadsheet GetSpreadsheet()
        {
            var spreadsheet = new Spreadsheet();

            var table = new Table(false);

            table.Columns.AddRange(new[]
            {
                new TableColumn("Example column 1"),
                new TableColumn("Example column 2"),
            });

            table.Rows = new List<TableRow>()
            {
                new TableRow()
                {
                    Cells =
                    {
                        new TableCell("First cell"),
                        new TableCell("Second cell")
                    }
                }
            };

            var workSheet = new Worksheet()
            {
                SheetName = "Example sheet",
                Table = table
            };

            spreadsheet.Worksheets.Add(workSheet);

            return spreadsheet;
        }
    }
}
