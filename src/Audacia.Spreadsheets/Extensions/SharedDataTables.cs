using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlCellFormat = DocumentFormat.OpenXml.Spreadsheet.CellFormat;

namespace Audacia.Spreadsheets.Extensions
{
    public static class SharedDataTables
    {
        public static CellStyle GetCellFormat(this SharedDataTable sharedData, CellStyle cellStyle)
        {
            return sharedData.CellFormats.SingleOrDefault(cf =>
                cf.TextColour == cellStyle.TextColour &&
                cf.BackgroundColour == cellStyle.BackgroundColour &&
                cf.BorderTop == cellStyle.BorderTop &&
                cf.BorderRight == cellStyle.BorderRight &&
                cf.BorderBottom == cellStyle.BorderBottom &&
                cf.BorderLeft == cellStyle.BorderLeft &&
                cf.Format == cellStyle.Format &&
                cf.HasWordWrap == cellStyle.HasWordWrap);
        }

        public static CellStyle GetOrCreateCellFormat(this SharedDataTable sharedData, CellStyle cellStyle)
        {
            var matchingStyle = sharedData.GetCellFormat(cellStyle);

            if (matchingStyle != default(CellStyle))
            {
                return matchingStyle;
            }

            var borders = new List<CellBorder>();
            if (cellStyle.BorderTop)
            {
                borders.Add(CellBorder.Top);
            }
            if (cellStyle.BorderRight)
            {
                borders.Add(CellBorder.Right);
            }
            if (cellStyle.BorderBottom)
            {
                borders.Add(CellBorder.Bottom);
            }
            if (cellStyle.BorderLeft)
            {
                borders.Add(CellBorder.Left);
            }

            var borderSum = UInt32Value.FromUInt32((uint) borders.Sum(b => (int) b));

            // Ignore custom boolean formats, see CellFormat.cs
            UInt32Value numberFormat = (uint) cellStyle.Format > 999U
                ? UInt32Value.FromUInt32((uint) CellFormat.Text)
                : UInt32Value.FromUInt32((uint) cellStyle.Format);

            var cellFormatsElement = sharedData.Stylesheet.CellFormats;

            var cellFormat = new OpenXmlCellFormat
            {
                FontId = cellStyle.TextColour,
                FillId = cellStyle.BackgroundColour,
                BorderId = borderSum,
                NumberFormatId = numberFormat,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true,
                ApplyNumberFormat = true
            };
            var alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Top,
                TextRotation = 0U,
                WrapText = cellStyle.HasWordWrap,
                ReadingOrder = 1U
            };

            cellFormat.Append(alignment);
            cellFormatsElement.Append(cellFormat);

            cellStyle.Index = UInt32Value.FromUInt32((uint) cellFormatsElement.ChildElements.Count) - 1;
            sharedData.CellFormats.Add(cellStyle);

            return cellStyle;
        }
    }
}