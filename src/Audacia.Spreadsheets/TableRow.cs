using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class TableRow
    {
        public int? Id { get; set; }
        public IList<TableCell> Cells { get; } = new List<TableCell>();

        public void Write(CellReference cellReference, IList<TableColumn> columns, SharedDataTable sharedData, OpenXmlWriter writer)
        {
            writer.WriteStartElement(new Row());
                
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var column = columns[columnIndex];
                var cell = Cells[columnIndex];
                var value = cell.Value;

                var cellStyle = new CellStyle
                {
                    TextColour = 0U,
                    BackgroundColour = 0U,
                    BorderBottom = true,
                    BorderTop = true,
                    BorderLeft = true,
                    BorderRight = true,
                    Format = column.Format,
                    HasWordWrap = value is string && !cell.IsFormula
                };

                if (!string.IsNullOrWhiteSpace(cell.FillColour))
                {
                    cellStyle.BackgroundColour = sharedData.FillColours[cell.FillColour];
                }

                if (!string.IsNullOrWhiteSpace(cell.TextColour))
                {
                    cellStyle.TextColour = sharedData.TextColours[cell.TextColour];
                }

                var styleIndex = sharedData.GetOrCreateCellFormat(cellStyle).Index;

                cell.Write(styleIndex, cellReference, writer);

                cellReference.NextColumn();
            }

            writer.WriteEndElement();
        }

        public static TableRow FromCells(IEnumerable<TableCell> cells, int? id)
        {
            var row = new TableRow
            {
                Id = id
            };
            
            foreach (var cell in cells)
            {
                row.Cells.Add(cell);
            }
            
            return row;
        }
        
        public static IEnumerable<TableRow> FromOpenXml(WorksheetPart worksheetPart, 
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
                var cellData = new List<TableCell>();

                for (var j = 0; j < columnsCount; j++)
                {
                    var cellReference = (j + 1).ToColumnLetter() + row.RowIndex;
                    var matchedCells =
                        cells.Where(
                            c =>
                                string.Compare(c.CellReference.Value, cellReference,
                                    StringComparison.OrdinalIgnoreCase) == 0).ToList();

                    if (!matchedCells.Any() || matchedCells.First().CellValue == null)
                    {
                        cellData.Add(new TableCell { Value = null });
                    }
                    else
                    {
                        var c = matchedCells.First();
                        if (c.DataType != null && c.DataType.HasValue && c.DataType.Value == CellValues.SharedString)
                        {
                            cellData.Add(new TableCell
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

                                        cellData.Add(new TableCell { Value = date });
                                        valueAdded = true;
                                    }
                                }
                            }

                            if (!valueAdded)
                            {
                                cellData.Add(new TableCell { Value = c.CellValue.Text });
                            }
                        }
                    }
                }

                if (cellData.All(c => c.Value == null || (c.Value is string s && string.IsNullOrWhiteSpace(s)))) continue;

                yield return TableRow.FromCells(cellData, rowNumber);
                
                rowNumber++;
            }
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
    }
}
