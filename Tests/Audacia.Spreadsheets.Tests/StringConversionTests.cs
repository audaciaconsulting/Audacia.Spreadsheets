using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using Audacia.Spreadsheets.Tests.Models.Unformatted;
using Audacia.Spreadsheets.Validation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Audacia.Spreadsheets.Tests
{
    /// <summary>
    /// Ensure that <see cref="WorksheetImporter{TRowModel}"/> can parse all types from general cell text that a user has entered.
    /// </summary>
    public class StringConversionTests
    {
        [Fact]
        public void BooleanConversions()
        {
            var expected = new[]
            {
                true,
                true,
                true,
                true,
                true,
                true,
                false,
                false,
                false,
                false,
                false,
                false
            };

            var actual = ConvertInputValues<BooleanModel>(new StringModel[]
            {
                "1",
                "y",
                "Y",
                "Yes",
                "true",
                "TRUE",
                "0",
                "n",
                "N",
                "No",
                "false",
                "FALSE"
            });

            ValidateAllRowsParsedCorrectly(expected, actual, a => a.Value);
        }

        [Fact]
        public void DateTimeConversions()
        {
            var expected = new[]
            {
                new DateTime(2021, 7, 20, 20, 43, 23),
                new DateTime(2020, 11, 30, 12, 0, 0),
                new DateTime(2018, 3, 2, 0, 0, 0),
                new DateTime(2016, 10, 1, 8, 35, 5),
                new DateTime(2011, 5, 1, 18, 30, 0),
                new DateTime(1970, 1, 1, 0, 0, 0)
            };

            var actual = ConvertInputValues<DateTimeModel>(new StringModel[]
            {
                "20/07/2021 20:43:23",
                "30/11/2020 12:00",
                "02/03/2018",
                "2016-10-01 08:35:05",
                "2011-05-01 18:30",
                "1970-01-01"
            });

            ValidateAllRowsParsedCorrectly(expected, actual, a => a.Value);
        }

        [Fact]
        public void DateTimeOffsetConversions()
        {
            var expected = new[]
            {
                new DateTimeOffset(2021, 7, 24, 10, 38, 00, TimeSpan.FromHours(-2)),
                new DateTimeOffset(2021, 7, 24, 11, 53, 00, TimeSpan.FromHours(5)),
                BuildExpectedWithLocalOffset(new DateTime(2021, 7, 20, 20, 43, 23)),
                new DateTimeOffset(2020, 11, 30, 12, 0, 0, TimeSpan.Zero),
                new DateTimeOffset(2018, 3, 2, 0, 0, 0, TimeSpan.Zero),
                BuildExpectedWithLocalOffset(new DateTime(2016, 10, 1, 8, 35, 5)),
                BuildExpectedWithLocalOffset(new DateTime(2011, 5, 1, 18, 30, 0)),
                new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero)
            };

            var actual = ConvertInputValues<DateTimeOffsetModel>(new StringModel[]
            {
                "24/07/2021 10:38:00 -02:00",
                "24/07/2021 11:53:00 +05:00",
                "20/07/2021 20:43:23",
                "30/11/2020 12:00",
                "02/03/2018",
                "2016-10-01 08:35:05",
                "2011-05-01 18:30",
                "1970-01-01"
            });

            ValidateAllRowsParsedCorrectly(expected, actual, a => a.Value);
        }

        [Fact]
        public void EnumConversions()
        {
            var expected = new[]
            {
                EnumModel.Shape.Hexagon,
                EnumModel.Shape.Pentagon,
                EnumModel.Shape.Square,
                EnumModel.Shape.Triangle,
                EnumModel.Shape.Circle,
                EnumModel.Shape.Pentagon,
                EnumModel.Shape.Square,
                EnumModel.Shape.Hexagon
            };

            var actual = ConvertInputValues<EnumModel>(new StringModel[]
            {
                "hexagon",
                "pentagon",
                "square",
                "triangle",
                "circle",
                "3",
                "2",
                "4"
            });

            ValidateAllRowsParsedCorrectly(expected, actual, a => a.Value);
        }

        [Fact]
        public void EnumOutOfRangeConversions()
        {
            var actual = ConvertInputValues<EnumModel>(new StringModel[]
            {
                "99",
                "pyramid"
            });

            ValidateUnableToParseRow(actual[0]);

            ValidateUnableToParseRow(actual[1]);
        }

        [Fact]
        public void TimeSpanConversions()
        {
            var expected = new[]
            {
                new TimeSpan(20, 43, 23),
                new TimeSpan(8, 35, 5),
                new TimeSpan(0, 0, 0),
                new TimeSpan(12, 0, 0),
                new TimeSpan(18, 30, 0),
                new TimeSpan(0, 0, 0)
            };

            var actual = ConvertInputValues<TimeSpanModel>(new StringModel[]
            {
                "20:43:23",
                "08:35:05",
                "00:00:00",
                "12:00",
                "18:30",
                "00:00"
            });

            ValidateAllRowsParsedCorrectly(expected, actual, a => a.Value);
        }

        /// <summary>
        /// Asserts that when parsing a value that Excel has converted to a bool, the string value is retained (i.e. it is not 1/0).
        /// </summary>
        [Fact]
        public void FromOpenXmlParsesBooleanCellsAsTrueOrFalseStrings()
        {
            using var spreadsheetStream = GenerateSpreadsheetStreamWithBooleanStrings();

            // Act: Parse the spreadsheet.
            using var readDoc = SpreadsheetDocument.Open(spreadsheetStream, false);
            var worksheetPartRead = readDoc.WorkbookPart!.WorksheetParts.First();
            var rows = TableRow.FromOpenXml(worksheetPartRead, readDoc, 1).ToList();

            // Assert: The boolean values are parsed as "TRUE" and "FALSE".
            Assert.Equal("TRUE", rows[1].Cells[0].Value);
            Assert.Equal("FALSE", rows[2].Cells[0].Value);
        }

        /// <summary>
        /// Converts the given <see cref="DateTime"/> to a <see cref="DateTimeOffset"/> offset by the local system time zone.
        /// </summary>
        /// <param name="expectedDateTime">The <see cref="DateTime"/> value to offset.</param>
        /// <returns>A <see cref="DateTimeOffset"/> object offset to the local system time zone.</returns>
        private static DateTimeOffset BuildExpectedWithLocalOffset(DateTime expectedDateTime)
        {
            var currentTimeZone = TimeZoneInfo.Local;
            var expectedOffset = currentTimeZone.GetUtcOffset(expectedDateTime);

            return new DateTimeOffset(expectedDateTime, expectedOffset);
        }

        /// <summary>
        /// Converts a string value to a typed value by converting to and from a spreadsheet.
        /// </summary>
        /// <typeparam name="T">Row Model</typeparam>
        /// <param name="source">Collection of input rows</param>
        /// <summary>
        private static IList<ImportRow<T>> ConvertInputValues<T>(IList<StringModel> source)
            where T : class, new()
        {
            // Export row models into spreadsheet file
            var bytes = Spreadsheet.FromWorksheets(source.ToWorksheet()).Export();

            // Read and parse spreadsheet into row models
            return new WorksheetImporter<T>()
                .ParseWorksheet(Spreadsheet.FromBytes(bytes).Worksheets[0])
                .ToArray();
        }

        /// <summary>
        /// Validates that all rows were parsed successfully as their expected values.
        /// </summary>
        /// <typeparam name="N">Expected Type</typeparam>
        /// <typeparam name="T">Expected Row Model</typeparam>
        /// <param name="expected">Expected values</param>
        /// <param name="output">Imported values</param>
        /// <param name="propertyFunc">Property to compare</param>
        private static void ValidateAllRowsParsedCorrectly<N, T>(IList<N> expected, IList<ImportRow<T>> output, Func<T, N> propertyFunc)
        {
            // Map sucessfully parsed rows, put a default value where invalid
            var actual = output
                .Select(importRow => importRow.IsValid
                    ? propertyFunc(importRow.Data)
                    : default(N)!)
                .ToArray();

            // Assert parsed collection matches the expected collection

            Assert.Equal(expected, actual);

            // Ensure that a parsing failure isn't ignored
            Assert.True(output.All(t => t.IsValid));
        }

        /// <summary>
        /// Validate that a single row was unable to be parsed.
        /// </summary>
        /// <typeparam name="T">Row Model</typeparam>
        /// <param name="output">Parsed row model</param>
        private static void ValidateUnableToParseRow<T>(ImportRow<T> output)
        {
            Assert.False(output.IsValid);

            Assert.Equal(1, output.ImportErrors.Count);

            Assert.IsType<FieldParseError>(output.ImportErrors.First());
        }

        /// <summary>
        /// Generates a sample spreadsheet with cells containing boolean strings (i.e. TRUE and FALSE), and returns it as a <see cref="Stream"/>.
        /// </summary>
        private static MemoryStream GenerateSpreadsheetStreamWithBooleanStrings()
        {
            var stream = new MemoryStream();

            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                var sharedStringTable = new SharedStringTable();
                sharedStringTable.Append(new SharedStringItem(new Text("Example Column")));

                sharedStringTable.Append(new SharedStringItem(new Text("TRUE")));
                sharedStringTable.Append(new SharedStringItem(new Text("FALSE")));

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

                var trueRow = new Row() { RowIndex = 2 };
                trueRow.Append(
                    new Cell
                    {
                        CellReference = "A2",
                        DataType = CellValues.Boolean,
                        CellValue = new CellValue("1")
                    });

                var falseRow = new Row() { RowIndex = 3 };
                falseRow.Append(
                    new Cell
                    {
                        CellReference = "A3",
                        DataType = CellValues.Boolean,
                        CellValue = new CellValue("0")
                    });

                sheetData.Append(trueRow);
                sheetData.Append(falseRow);

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