using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlCellFormat = DocumentFormat.OpenXml.Spreadsheet.CellFormat;

namespace Audacia.Spreadsheets
{
    public class TableRow
    {
        public int? Id { get; set; }
        public List<TableCell> Cells { get; } = new List<TableCell>();

        public void Write(CellReference cellReference, IList<TableColumn> columns, SharedDataTable sharedData, OpenXmlWriter writer)
        {
            writer.WriteStartElement(new Row());
                
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var column = columns[columnIndex];
                var cell = Cells.Count > columnIndex
                    ? Cells[columnIndex]
                    : new TableCell(hasBorders: Cells.Count > 0 && Cells.Last().HasBorders);
                
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
                    HasWordWrap = value is string && !cell.IsFormula,
                    IsEditable = cell.IsEditable
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

                cell.Write(styleIndex, column.Format, cellReference, writer);

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
            SpreadsheetDocument spreadSheet, int columnsCount, int startingRowIndex = 0)
        {
            var stylesheet = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
            var cellFormats = stylesheet.CellFormats;
            var dateFormatIds = GetDateFormatsInFile(spreadSheet.WorkbookPart.WorkbookStylesPart);

            // Read each row and add to table
            var rows = worksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().ToList();
            var stringTable = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();

            var rowPointer = new CellReference("A1").MutateBy(0, startingRowIndex);
            
            foreach (var row in rows.Skip(startingRowIndex))
            {
                var cellRef = rowPointer.Clone();
                var cells = row.Elements<Cell>().ToArray();
                var cellData = new List<TableCell>();
                
                for (var j = 0; j < columnsCount; j++)
                {
                    var cellReference = cellRef.ToString();
                    var matchedCells = cells.Where(c =>
                        string.Compare(cellReference, c.CellReference.Value, StringComparison.OrdinalIgnoreCase) == 0)
                        .ToList();
                    var newCell = new TableCell(null);

                    if (!matchedCells.Any() || matchedCells.First().CellValue == null)
                    {
                        cellData.Add(newCell);
                    }
                    else
                    {
                        var c = matchedCells.First();
                        if (c.DataType != null && c.DataType.HasValue && c.DataType.Value == CellValues.SharedString)
                        {
                            newCell.Value = stringTable.SharedStringTable.ElementAt(int.Parse(c.CellValue.Text)).InnerText;

                            if (c.StyleIndex != null)
                            {
                                var styleIndex = (int)c.StyleIndex.Value;
                                var cellFormat = (OpenXmlCellFormat)cellFormats.ElementAt(styleIndex);

                                var fill = (Fill)stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value];
                                var patternFill = fill?.PatternFill;

                                if (patternFill != null)
                                {
                                    newCell.FillColour = GetColor(spreadSheet, patternFill);
                                }
                            }

                            cellData.Add(newCell);
                        }
                        else
                        {
                            var valueAdded = false;

                            if (c.StyleIndex != null)
                            {
                                var styleIndex = (int)c.StyleIndex.Value;
                                var cellFormat = (OpenXmlCellFormat)cellFormats.ElementAt(styleIndex);

                                Fill fill = (Fill)stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value];
                                var patternFill = fill?.PatternFill;

                                if (IsDateFormat(cellFormat.NumberFormatId) ||
                                    dateFormatIds.Contains(cellFormat.NumberFormatId))
                                {
                                    if (double.TryParse(c.CellValue.InnerXml, out var parsedValue))
                                    {
                                        var date = DateTime.FromOADate(parsedValue);

                                        newCell.Value = date;
                                        cellData.Add(newCell);
                                        valueAdded = true;
                                    }
                                }

                                if (!valueAdded && IsNumberFormat(cellFormat.NumberFormatId))
                                {
                                    if (!valueAdded && decimal.TryParse(c.CellValue.Text, out var value))
                                    {
                                        newCell.Value = value;
                                        cellData.Add(newCell);
                                        valueAdded = true;
                                    }
                                }

                                newCell.FillColour = GetColor(spreadSheet, patternFill);
                            }

                            if (!valueAdded)
                            {
                                newCell.Value = c.CellValue.Text;
                                cellData.Add(newCell);
                            }


                        }
                    }
                    
                    cellRef.NextColumn();
                }

                if (cellData.All(c => c.Value == null || (c.Value is string s && string.IsNullOrWhiteSpace(s)))) continue;

                var rowId = Convert.ToInt32(rowPointer.RowNumber);
                yield return FromCells(cellData, rowId);
                
                rowPointer.NextRow();
            }
        }

        private static string GetColor(SpreadsheetDocument sd, PatternFill fill)
        {
            ColorType ct = fill.ForegroundColor;
            if (ct == null)
            {
                return null;
            }

            if (ct.Auto != null)
            {
                return null;
            }

            if (ct.Rgb != null)
            {
                //  ct.Rgb gives "FF######" so need to take off the first 2 characters. Thanks OpenXml
                return ct.Rgb.Value.Substring(2);
            }

            //  These 3 are too difficult to understand...Code stolen from:
            //  https://stackoverflow.com/questions/10756206/getting-cell-backgroundcolor-in-excel-with-open-xml-2-0
            //  
            //if (ct.Indexed != null)
            //{
            //    return sd.WorkbookPart.WorkbookStylesPart.Stylesheet.Colors.IndexedColors.ChildElements[(int)ct.Indexed.Value].InnerText;
            //}

            //if (ct.Theme != null)
            //{
            //    var c2t = (DocumentFormat.OpenXml.Drawing.Color2Type)sd.WorkbookPart.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)ct.Theme.Value];

            //    return ((DocumentFormat.OpenXml.Drawing.SystemColor)c2t.FirstChild).LastColor;
            //}

            //if (ct.Tint != null)
            //{
            //    return ct.Tint.Value.ToString();
            //}

            return null;
        }


        private static bool IsDateFormat(uint numberFormatId, string formatCode = null)
        {
            // Microsoft only give limited format information, there's no entire list of format codes online
            // So first check the ones we do
            if ((numberFormatId >= (uint)CellFormat.Date 
              && numberFormatId <= (uint)CellFormat.DateTime) 
              || numberFormatId == 30U)
            {
                return true;
            }

            // Then check if it contains date formatting or year formatting
            return formatCode != null && (formatCode.Contains("mmm") || formatCode.Contains("yy"));
        }

        private static bool IsNumberFormat(uint numberFormatId, string formatCode = null)
        {
            // Microsoft only give limited format information, there's no entire list of format codes online
            // So first check the ones we do
            if ((numberFormatId >= (uint)CellFormat.Integer 
              && numberFormatId <= (uint)CellFormat.Scientific)
              || numberFormatId == (uint)CellFormat.Currency
              || (numberFormatId >= (uint)CellFormat.AccountingGBP 
              && numberFormatId <= (uint)CellFormat.AccountingEUR))
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
