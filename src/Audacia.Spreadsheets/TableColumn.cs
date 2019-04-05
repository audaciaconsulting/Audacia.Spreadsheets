﻿using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Attributes;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class TableColumn
    {
        public TableColumn() { }

        public TableColumn(string name)
        {
            Name = name;
        }

        public TableColumn(string name, CellFormat format)
        {
            Name = name;
            Format = format;
        }
        
        public TableColumn(string name, CellFormat format, bool displaySubtotal)
        {
            Name = name;
            Format = format;
            DisplaySubtotal = displaySubtotal;
        }
        
        public TableColumn(string name, CellFormat format = CellFormat.Text, bool displaySubtotal = false, bool hasBorders = true)
        {
            Name = name;
            Format = format;
            DisplaySubtotal = displaySubtotal;
            HasBorders = hasBorders;
        }

        public string Name { get; set; }

        public bool DisplaySubtotal { get; set; }

        public int? Width { get; set; }

        public bool HasBorders { get; set; } = true;
        
        public CellFormat Format { get; set; } = CellFormat.Text;

        public CellBackgroundColourAttribute CellBackgroundFormat { get; set; }

        public CellTextColourAttribute CellTextFormat { get; set; }

        /// <summary>
        /// Writes a subtotal formulae above the current column header.
        /// </summary>
        public void WriteSubtotal(CellReference cellReference, bool isFirstColumn, bool isLastColumn, int totalRows, SharedDataTable sharedData, OpenXmlWriter writer)
        {
            var cellStyle = new CellStyle
            {
                TextColour = 1U,
                BackgroundColour = 0U,
                BorderBottom = HasBorders,
                BorderTop = HasBorders,
                BorderLeft = HasBorders && isFirstColumn,
                BorderRight = HasBorders && isLastColumn,
                Format = DisplaySubtotal ? Format : CellFormat.Text,
                HasWordWrap = false
            };

            var styleIndex = sharedData.GetOrCreateCellFormat(cellStyle).Index;
            var dataType = DataType.String;
            var formula = string.Empty;
            
            if (DisplaySubtotal)
            {
                // Increment by 2 so that the formula starts after the header row & the current row
                var firstRow = cellReference.MutateBy(0, 2);

                // Doesn't need to include first row
                // because the formulae starts on the first row of data
                var totalRowsAfterFirst = totalRows == 0 ? 0 : totalRows - 1;
                var lastRow = firstRow.MutateBy(0, totalRowsAfterFirst);

                // If we use SUBTOTAL(9,XX:XX) then it recalculates as the filter changes...
                formula = $"SUBTOTAL(9,{firstRow}:{lastRow})";
                dataType = DataType.Number;
            }
            
            TableCell.WriteCell(writer, styleIndex, cellReference, dataType, formula, DisplaySubtotal);
        }
        
        /// <summary>
        /// Writes the current column header.
        /// </summary>
        public void Write(TableHeaderStyle headerStyle, CellReference cellReference, bool isFirstColumn, bool isLastColumn, SharedDataTable sharedData, OpenXmlWriter writer)
        {
            var noHeaderStyle = headerStyle == default(TableHeaderStyle);
            
            if (noHeaderStyle || !sharedData.Fonts.TryGetValue($"{headerStyle.FontName}:{headerStyle.TextColour}", out var font))
            {
                font = 1u;
            }

            if (noHeaderStyle || !sharedData.FillColours.TryGetValue(headerStyle.FillColour, out var fillColour))
            {
                fillColour = 2u;
            }

            var cellStyle = new CellStyle
            {
                TextColour = font,
                BackgroundColour = fillColour,
                BorderBottom = HasBorders,
                BorderTop = HasBorders,
                BorderLeft = HasBorders && isFirstColumn,
                BorderRight = HasBorders && isLastColumn,
                Format = CellFormat.Text,
                HasWordWrap = false
            };

            var styleIndex = sharedData.GetOrCreateCellFormat(cellStyle).Index;

            TableCell.WriteCell(writer, styleIndex, cellReference, DataType.String, Name, false);
        }
        
        public static IEnumerable<TableColumn> FromOpenXml(WorksheetPart worksheetPart, SpreadsheetDocument spreadSheet, bool hasSubtotals)
        {
            // Get column headers
            var cellReference = new CellReference("A1");
            
            if (hasSubtotals) cellReference.NextRow();
            
            string newHeader;
            do
            {
                newHeader = GetColumnHeading(spreadSheet, worksheetPart, cellReference);
                if (!string.IsNullOrWhiteSpace(newHeader))
                {
                    yield return new TableColumn(newHeader.Trim());
                }
                cellReference.NextColumn();
            } while (!string.IsNullOrWhiteSpace(newHeader));
        }
        
        // Given a document name, a worksheet name, and a cell name, gets the column of the cell and returns
        // the content of the first cell in that column.
        private static string GetColumnHeading(SpreadsheetDocument document, WorksheetPart worksheetPart,
            CellReference cellReference)
        {
            // Get the column name for the specified cell.
            var columnName = cellReference.ColumnLetter;

            // Get the cells in the specified column and order them by row.
            var cells = worksheetPart.Worksheet.Descendants<Cell>()
                .Select(cell => (new CellReference(cell.CellReference.Value), cell))
                .Where(c =>
                {
                    var columnLetter = c.Item1.ColumnLetter;
                    var isInColumn = string.Compare(columnLetter, columnName, StringComparison.OrdinalIgnoreCase) == 0;
                    var isInRow = c.Item1.RowNumber == cellReference.RowNumber;
                    return isInColumn && isInRow;
                })
                .Select(c => c.Item2);

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
    }
}