using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets.Extensions
{
    public static class SharedDataTables
    {
        public static CellStyle GetOrCreateCellFormat(this SharedDataTable sharedData, CellStyle cellStyle)
        {
            var matchingStyle = sharedData.CellFormats.SingleOrDefault(cf =>
                cf.TextColour == cellStyle.TextColour &&
                cf.BackgroundColour == cellStyle.BackgroundColour &&
                cf.BorderTop == cellStyle.BorderTop &&
                cf.BorderRight == cellStyle.BorderRight &&
                cf.BorderBottom == cellStyle.BorderBottom &&
                cf.BorderLeft == cellStyle.BorderLeft &&
                cf.Format == cellStyle.Format &&
                cf.HasWordWrap == cellStyle.HasWordWrap);

            if (matchingStyle != default(CellStyle))
            {
                return matchingStyle;
            }

            var borders = new List<CellBorderType>();
            if (cellStyle.BorderTop) borders.Add(CellBorderType.Top);
            if (cellStyle.BorderRight) borders.Add(CellBorderType.Right);
            if (cellStyle.BorderBottom) borders.Add(CellBorderType.Bottom);
            if (cellStyle.BorderLeft) borders.Add(CellBorderType.Left);

            var borderSum = UInt32Value.FromUInt32((uint) borders.Sum(b => (int) b));

            UInt32Value numberFormat;
            switch (cellStyle.Format)
            {
                case CellFormatType.Date:
                    numberFormat = 14U;
                    break;
                case CellFormatType.Currency:
                    numberFormat = 165U;
                    break;
                // ReSharper disable once RedundantCaseLabel
                case CellFormatType.Text:
                default:
                    numberFormat = 0U;
                    break;
            }

            var cellFormatsElement = sharedData.Stylesheet.CellFormats;

            var cellFormat = new CellFormat
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