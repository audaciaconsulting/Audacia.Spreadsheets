using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    // TODO JP: come back and optimise this
    public class SpreadsheetParser
    {
        public Spreadsheet GetSpreadsheetFromExcelFile(Stream stream, bool includeHeaders = true)
        {
            using (var spreadSheet = SpreadsheetDocument.Open(stream, false))
            {
                var worksheets = spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>()
                    .Select((sheet, index) => ReturnDataModel(sheet, spreadSheet, index, includeHeaders))
                    .ToList();
                return Spreadsheet.FromWorksheets(worksheets);
            }
        }

        private static Worksheet ReturnDataModel(Sheet worksheet, SpreadsheetDocument spreadSheet, int index,
            bool includeHeaders)
        {
            var worksheetPart = (WorksheetPart) spreadSheet.WorkbookPart.GetPartById(worksheet.Id);

            var table = new Table
            {
                StartingCellRef = "A1",
                IncludeHeaders = includeHeaders,
                HeaderStyle = null
            };

            if (includeHeaders)
            {
                var columns = GetColumnModels(worksheetPart, spreadSheet);
                table.Columns.AddRange(columns);
            }

            var maxRowWidth = includeHeaders ? table.Columns.Count : GetMaxRowWidth(worksheetPart);

            var rows = GetRowModels(worksheetPart, spreadSheet, maxRowWidth, includeHeaders);
            table.Rows.AddRange(rows);

            return new Worksheet
            {
                SheetName = worksheet.Name,
                SheetIndex = index,
                Tables = new List<Table> { table }
            };
        }

        private static IEnumerable<WorksheetTableColumn> GetColumnModels(WorksheetPart worksheetPart, 
            SpreadsheetDocument spreadSheet)
        {
            // Get column headers
            var i = 1;
            string newHeader;
            do
            {
                var columnName = i.GetExcelColumnName();
                newHeader = GetColumnHeading(spreadSheet, worksheetPart, columnName + "1");
                if (!string.IsNullOrWhiteSpace(newHeader))
                {
                    yield return new WorksheetTableColumn { Name = newHeader };
                }
                i++;
            } while (!string.IsNullOrWhiteSpace(newHeader));
        }

        private static IEnumerable<WorksheetTableRow> GetRowModels(WorksheetPart worksheetPart, 
            SpreadsheetDocument spreadSheet, int columnsCount, bool includeHeaders)
        {
            var cellFormats = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
            var dateFormatIds = GetDateFormatsInFile(spreadSheet.WorkbookPart.WorkbookStylesPart);

            // Read each row and add to table
            var rows = worksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().ToList();
            var stringTable = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();

            var rowNumber = 1;

            // Starts at i = 1 to skip header row IF headers are included
            foreach (var row in rows.Skip(includeHeaders ? 1 : 0))
            {
                var cells = row.Elements<Cell>().ToArray();
                var cellData = new List<WorksheetTableCell>();

                for (var j = 0; j < columnsCount; j++)
                {
                    var cellReference = (j + 1).GetExcelColumnName() + row.RowIndex;
                    var matchedCells =
                        cells.Where(
                            c =>
                                string.Compare(c.CellReference.Value, cellReference,
                                    StringComparison.OrdinalIgnoreCase) == 0).ToList();

                    if (!matchedCells.Any() || matchedCells.First().CellValue == null)
                    {
                        cellData.Add(new WorksheetTableCell { Value = null });
                    }
                    else
                    {
                        var c = matchedCells.First();
                        if (c.DataType != null && c.DataType.HasValue && c.DataType.Value == CellValues.SharedString)
                        {
                            cellData.Add(new WorksheetTableCell
                            {
                                Value = stringTable.SharedStringTable.ElementAt(int.Parse(c.CellValue.Text)).InnerText
                            });
                        }
                        else
                        {
                            var valueAdded = false;

                            if (c.StyleIndex != null)
                            {
                                var styleIndex = (int)c.StyleIndex.Value;
                                var cellFormat = (CellFormat)cellFormats.ElementAt(styleIndex);

                                if (IsDateFormat(cellFormat.NumberFormatId) ||
                                    dateFormatIds.Contains(cellFormat.NumberFormatId))
                                {
                                    if (double.TryParse(c.CellValue.InnerXml, out var parsedValue))
                                    {
                                        var date = DateTime.FromOADate(parsedValue);

                                        cellData.Add(new WorksheetTableCell { Value = date });
                                        valueAdded = true;
                                    }
                                }
                            }

                            if (!valueAdded)
                            {
                                cellData.Add(new WorksheetTableCell { Value = c.CellValue.Text });
                            }
                        }
                    }
                }

                if (cellData.All(c => c.Value == null || (c.Value is string s && string.IsNullOrWhiteSpace(s)))) continue;

                yield return WorksheetTableRow.FromCells(cellData, rowNumber);
                
                rowNumber++;
            }
        }

        // Given a document name, a worksheet name, and a cell name, gets the column of the cell and returns
        // the content of the first cell in that column.
        private static string GetColumnHeading(SpreadsheetDocument document, WorksheetPart worksheetPart,
            string cellName)
        {
            // Get the column name for the specified cell.
            var columnName = GetColumnName(cellName);

            // Get the cells in the specified column and order them by row.
            var cells = worksheetPart.Worksheet.Descendants<Cell>()
                .Where(
                    c =>
                        string.Compare(GetColumnName(c.CellReference.Value), columnName,
                            StringComparison.OrdinalIgnoreCase) == 0)
                .OrderBy(r => GetRowIndex(r.CellReference));

            // Get the first cell in the column.
            var headCell = cells.FirstOrDefault();

            if (headCell == default(Cell))
            {
                // The specified column does not exist.
                return null;
            }

            // If the content of the first cell is stored as a shared string, get the text of the first cell
            // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            if (headCell.DataType == null || headCell.DataType.Value != CellValues.SharedString)
            {
                return headCell.CellValue == null ? string.Empty : headCell.CellValue.Text;
            }

            var shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            var items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            return items[int.Parse(headCell.CellValue.Text)].InnerText;
        }

        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellName);

            return match.Value;
        }

        // Given a cell name, parses the specified cell to get the row index.
        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            var regex = new Regex(@"\d+");
            var match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        private static ICollection<uint> GetDateFormatsInFile(WorkbookStylesPart stylePart)
        {
            var formatIds = new Collection<uint>();

            var numFormatsParentNodes = stylePart.Stylesheet.ChildElements.OfType<NumberingFormats>();

            foreach (var numFormatParentNode in numFormatsParentNodes)
            {
                var formatNodes = numFormatParentNode.ChildElements.OfType<NumberingFormat>();
                foreach (var formatNode in formatNodes)
                {
                    if (IsDateFormat(formatNode.NumberFormatId.Value, formatNode.FormatCode))
                    {
                        formatIds.Add(formatNode.NumberFormatId.Value);
                    }
                }
            }

            return formatIds;
        }

        private static bool IsDateFormat(uint numberFormatId, string formatCode = null)
        {
            // Microsoft only give limited format information, there's no entire list of format codes online
            // So first check the ones we do
            if ((numberFormatId >= 14 && numberFormatId <= 22) || numberFormatId == 30)
            {
                return true;
            }

            // Then check if it contains date formatting or year formatting
            return formatCode != null && (formatCode.Contains("mmm") || formatCode.Contains("yy"));
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

                var rowIndex = GetRowIndex(lastCell.CellReference);
                if (rowIndex > maxWidth)
                {
                    maxWidth = (int)rowIndex;
                }
            }

            return maxWidth;
        }
    }
}
