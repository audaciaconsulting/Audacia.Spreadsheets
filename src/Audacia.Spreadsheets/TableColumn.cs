using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Attributes;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class WorksheetTableColumn
    {
        public WorksheetTableColumn() { }

        public WorksheetTableColumn(string name) => Name = name;

        public static implicit operator WorksheetTableColumn(string name) => new WorksheetTableColumn(name);

        public string Name { get; set; }

        public bool DisplaySubtotal { get; set; }

        public CellFormatType Format { get; set; } = CellFormatType.Text;

        public CellBackgroundColourAttribute CellBackgroundFormat { get; set; }

        public CellTextColourAttribute CellTextFormat { get; set; }

        /// <summary>
        /// Writes a subtotal formulae above the current column header.
        /// </summary>
        public void WriteSubtotal(CellReference cellReference, bool isFirstColumn, bool isLastColumn, int totalRows, SharedData sharedData, OpenXmlWriter writer)
        {
            var cellStyle = new CellStyle
            {
                TextColour = 1U,
                BackgroundColour = 0U,
                BorderBottom = true,
                BorderTop = true,
                BorderLeft = isFirstColumn,
                BorderRight = isLastColumn,
                Format = Format,
                HasWordWrap = false
            };

            var styleIndex = sharedData.GetOrCreateCellFormat(cellStyle).Index;
            var dataType = DataType.String;
            var formula = string.Empty;
            
            if (DisplaySubtotal)
            {
                // Increment by 2 so that the formula starts after the header row & the current row
                var firstRow = cellReference.MutateRowsBy(2);

                // Doesn't need to include first row
                // because the formulae starts on the first row of data
                var totalRowsAfterFirst = totalRows == 0 ? 0 : totalRows - 1;
                var lastRow = firstRow.MutateRowsBy(totalRowsAfterFirst);

                // If we use SUBTOTAL(9,XX:XX) then it recalculates as the filter changes...
                formula = $"SUBTOTAL(9,{firstRow}:{lastRow})";
                dataType = DataType.Numeric;
            }
            
            TableCell.WriteCell(writer, styleIndex, cellReference, dataType, formula, DisplaySubtotal);
        }
        
        /// <summary>
        /// Writes the current column header
        /// </summary>
        public void Write(TableHeaderStyle headerStyle, CellReference cellReference, bool isFirstColumn, bool isLastColumn, SharedData sharedData, OpenXmlWriter writer)
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
                BorderBottom = true,
                BorderTop = true,
                BorderLeft = isFirstColumn,
                BorderRight = isLastColumn,
                Format = CellFormatType.Text,
                HasWordWrap = false
            };

            var styleIndex = sharedData.GetOrCreateCellFormat(cellStyle).Index;

            TableCell.WriteCell(writer, styleIndex, cellReference, DataType.String, Name, false);
        }
    }
}
