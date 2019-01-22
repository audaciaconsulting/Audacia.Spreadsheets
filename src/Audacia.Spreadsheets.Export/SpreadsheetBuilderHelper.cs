﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using Audacia.Spreadsheets.Models.Constants;
using Audacia.Spreadsheets.Models.Enums;
using Audacia.Spreadsheets.Models.WorksheetData;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets.Export
{
    public static class SpreadsheetBuilderHelper
    {
        public const string DefaultStartingCellRef = "A1";
        public static int MaxColumnWidth = 75;

        public static void Insert(TableModel model, Stylesheet stylesheet,
            List<SpreadsheetCellStyle> cellFormats, Dictionary<string, uint> fillColours,
            Dictionary<string, uint> textColours, Dictionary<string, uint> fonts, WorksheetPart worksheet,
            OpenXmlWriter writer)
        {
            var startCellReference = !string.IsNullOrWhiteSpace(model.StartingCellRef)
                ? model.StartingCellRef
                : DefaultStartingCellRef;

            var cellReferenceRowIndex = startCellReference.GetReferenceRowIndex();
            var cellReferenceColumnIndex = startCellReference.GetReferenceColumnIndex();

            if (model.IncludeHeaders)
            {
                writer.WriteStartElement(new Row());

                foreach (var column in model.Data.Columns)
                {
                    if (!fonts.TryGetValue($"{model.HeaderStyle?.FontName}:{model.HeaderStyle?.TextColour}", out var font))
                    {
                        font = 1u;
                    }

                    if (!fillColours.TryGetValue(model.HeaderStyle?.FillColour, out var fillColour))
                    {
                        fillColour = 2u;
                    }

                    var cellStyle = new SpreadsheetCellStyle
                    {
                        TextColour = font,
                        BackgroundColour = fillColour,
                        BorderBottom = true,
                        BorderTop = true,
                        BorderLeft = column == model.Data.Columns.ElementAt(0),
                        BorderRight = column == model.Data.Columns.ElementAt(model.Data.Columns.Count() - 1),
                        Format = CellFormatType.Text,
                        HasWordWrap = false
                    };

                    var styleIndex = GetOrCreateCellFormat(cellStyle, cellFormats, stylesheet).Index;

                    WriteCell(writer, styleIndex, $"{cellReferenceColumnIndex}{cellReferenceRowIndex}",
                        OpenXmlDataType.OpenXmlStringDataType, column.HideHeader ? string.Empty : column.Name);

                    //Update column reference for next iteration
                    cellReferenceColumnIndex = (cellReferenceColumnIndex.GetColumnNumber() + 1)
                        .GetExcelColumnName();
                }

                cellReferenceColumnIndex = startCellReference.GetReferenceColumnIndex();
                startCellReference = $"{cellReferenceColumnIndex}{cellReferenceRowIndex++}";

                writer.WriteEndElement();
            }

            foreach (var row in model.Data.Rows)
            {
                writer.WriteStartElement(new Row());
                var columnIndex = 0;
                foreach (var column in model.Data.Columns)
                {
                    var cellModel = row.Cells.ElementAt(columnIndex);
                    var value = cellModel.Value;

                    var cellStyle = new SpreadsheetCellStyle
                    {
                        TextColour = 0U,
                        BackgroundColour = 0U,
                        BorderBottom = true,
                        BorderTop = true,
                        BorderLeft = true,
                        BorderRight = true,
                        Format = value is DateTime ? CellFormatType.Date : column.Format,
                        HasWordWrap = value is string
                    };

                    if (!string.IsNullOrWhiteSpace(cellModel.FillColour))
                    {
                        cellStyle.BackgroundColour = fillColours[cellModel.FillColour];
                    }

                    if (!string.IsNullOrWhiteSpace(cellModel.TextColour))
                    {
                        cellStyle.TextColour = textColours[cellModel.TextColour];
                    }

                    var styleIndex = GetOrCreateCellFormat(cellStyle, cellFormats, stylesheet).Index;

                    var dataTypeAndValue = GetDataTypeAndFormattedValue(value);

                    WriteCell(writer, styleIndex, $"{cellReferenceColumnIndex}{cellReferenceRowIndex}",
                        dataTypeAndValue.Item1, dataTypeAndValue.Item2, cellModel.IsFormula);

                    cellReferenceColumnIndex = (cellReferenceColumnIndex.GetColumnNumber() + 1)
                        .GetExcelColumnName();

                    columnIndex++;
                }
                cellReferenceColumnIndex = startCellReference.GetReferenceColumnIndex();
                cellReferenceRowIndex++;
                writer.WriteEndElement();
            }
        }

        internal static void AddProtection(WorksheetPart worksheetPart, WorksheetProtection worksheetProtection)
        {
            var sheetProtection = new SheetProtection
            {
                Objects = true,
                Scenarios = true,
                Sheet = true,
                InsertColumns = !worksheetProtection.CanAddOrDeleteColumns,
                DeleteColumns = !worksheetProtection.CanAddOrDeleteColumns,
                InsertRows = !worksheetProtection.CanAddOrDeleteRows,
                DeleteRows = !worksheetProtection.CanAddOrDeleteRows,
            };

            if (!string.IsNullOrWhiteSpace(worksheetProtection.Password))
            {
                // NOTE: We cannot use Workbook protection, as the resulting OpenXML file is marked as corrupted
                // by OpenXML when attempting to open it - the Productivity tool does the same thing.
                // So we'll just do worksheet protection
                sheetProtection.Password = HexPasswordConversion(worksheetProtection.Password);
            }

            var pRanges = new ProtectedRanges();

            foreach (var protectedRange in worksheetProtection.EditableCellRanges)
            {
                var pRange = new ProtectedRange();
                var lValue = new ListValue<StringValue> { InnerText = protectedRange };

                pRange.SequenceOfReferences = lValue;
                pRange.Name = "not allow editing";
                pRanges.Append(pRange);
            }

            //These are the cells that are editable
            var pageM = worksheetPart.Worksheet.GetFirstChild<PageMargins>();
            worksheetPart.Worksheet.InsertBefore(sheetProtection, pageM);
            worksheetPart.Worksheet.InsertBefore(pRanges, pageM);
        }

        private static void WriteCell(OpenXmlWriter writer, UInt32Value styleIndex,
            string reference, string dataType, string value, bool isFormula = false)
        {
            var attributes = new List<OpenXmlAttribute>
            {
                new OpenXmlAttribute("r", null, reference),
                new OpenXmlAttribute("s", null, styleIndex),
                new OpenXmlAttribute("t", null, dataType)
            };

            writer.WriteStartElement(new Cell(), attributes);

            if (!string.IsNullOrWhiteSpace(value))
            {
                if (isFormula)
                {
                    writer.WriteElement(new CellFormula(value));
                }
                else
                {
                    writer.WriteElement(new CellValue(value));
                }
            }

            writer.WriteEndElement();
        }

        private static Tuple<string, string> GetDataTypeAndFormattedValue(object cellValue)
        {
            switch (cellValue)
            {
                case DateTime time:
                    return new Tuple<string, string>(
                        OpenXmlDataType.OpenXmlDateDataType, FormatDate(time));
                case decimal @decimal:
                    return new Tuple<string, string>(
                        OpenXmlDataType.OpenXmlNumericDataType, @decimal.ToString(CultureInfo.CurrentCulture));
                case double d:
                    return new Tuple<string, string>(
                        OpenXmlDataType.OpenXmlNumericDataType, d.ToString(CultureInfo.CurrentCulture));
                case int i:
                    return new Tuple<string, string>(
                        OpenXmlDataType.OpenXmlNumericDataType, i.ToString(CultureInfo.CurrentCulture));
                default:
                    return new Tuple<string, string>(
                                    OpenXmlDataType.OpenXmlStringDataType, cellValue.ToString());
            }
        }

        private static SpreadsheetCellStyle GetOrCreateCellFormat(SpreadsheetCellStyle cellStyle,
            ICollection<SpreadsheetCellStyle> cellFormats, Stylesheet stylesheet)
        {
            var matchingStyle = cellFormats.SingleOrDefault(cf => cf.TextColour == cellStyle.TextColour &&
                                                                  cf.BackgroundColour == cellStyle.BackgroundColour &&
                                                                  cf.BorderTop == cellStyle.BorderTop &&
                                                                  cf.BorderRight == cellStyle.BorderRight &&
                                                                  cf.BorderBottom == cellStyle.BorderBottom &&
                                                                  cf.BorderLeft == cellStyle.BorderLeft &&
                                                                  cf.Format == cellStyle.Format &&
                                                                  cf.HasWordWrap == cellStyle.HasWordWrap);

            if (matchingStyle != default(SpreadsheetCellStyle))
            {
                return matchingStyle;
            }

            var borders = new List<CellBorderType>();
            if (cellStyle.BorderTop) borders.Add(CellBorderType.Top);
            if (cellStyle.BorderRight) borders.Add(CellBorderType.Right);
            if (cellStyle.BorderBottom) borders.Add(CellBorderType.Bottom);
            if (cellStyle.BorderLeft) borders.Add(CellBorderType.Left);

            var borderSum = UInt32Value.FromUInt32((uint)borders.Sum(b => (int)b));

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

            var cellFormatsElement = stylesheet.CellFormats;

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

            cellStyle.Index = UInt32Value.FromUInt32((uint)cellFormatsElement.ChildElements.Count) - 1;
            cellFormats.Add(cellStyle);

            return cellStyle;
        }

        private static string FormatDate(DateTime value)
        {
            if (value.Equals(DateTime.MinValue))
            {
                return string.Empty;
            }

            return value.ToOADatePrecise().ToString(new CultureInfo("en-US"));
        }

        public static void AddSheetView(OpenXmlWriter writer)
        {
            writer.WriteStartElement(new SheetViews());
            writer.WriteElement(new SheetView
            {
                ShowGridLines = false,
                WorkbookViewId = 0U
            });
            writer.WriteEndElement();
        }

        public static void AddColumns(OpenXmlWriter writer, TableModel tableModel)
        {
            writer.WriteStartElement(new Columns());

            var maxColWidth = GetMaxCharacterWidth(tableModel.Data);
            double maxWidth = 11;

            for (var i = 0; i < maxColWidth.Count; i++)
            {
                var item = maxColWidth[i];

                //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {20 pixel padding}]/{Maximum Digit Width}*256)/256
                var width = Math.Truncate((item * maxWidth + 20) / maxWidth * 256) / 256;

                if (width > 75)
                {
                    width = 75;
                }

                var colWidth = (DoubleValue)width;

                writer.WriteElement(new Column
                {
                    Min = Convert.ToUInt32(i + 1),
                    Max = Convert.ToUInt32(i + 1),
                    CustomWidth = true,
                    BestFit = true,
                    Width = colWidth
                });
            }

            writer.WriteEndElement();
        }

        private static Dictionary<int, int> GetMaxCharacterWidth(TableWrapperModel model)
        {
            //iterate over all cells getting a max char value for each column
            var maxColWidth = new Dictionary<int, int>();

            var columnHeaderWithData = model.Rows.ToList();

            columnHeaderWithData.Add(
                new TableRowModel
                {
                    Cells = model.Columns.Select(c =>
                        new TableCellModel
                        {
                            Value = c.Name
                        })
                });

            foreach (var r in columnHeaderWithData)
            {
                var cells = r.Cells.ToArray();

                //using cell index as my column
                for (var i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.Value?.ToString() ?? string.Empty;
                    var cellTextLength = cellValue.Length;

                    if (maxColWidth.ContainsKey(i))
                    {
                        var current = maxColWidth[i];
                        if (cellTextLength > current)
                        {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }

        private static HexBinaryValue HexPasswordConversion(string password)
        {
            if (string.IsNullOrWhiteSpace(password))
            {
                throw new ArgumentException("Cannot convert an empty password");
            }

            byte[] passwordCharacters = System.Text.Encoding.ASCII.GetBytes(password);
            int hash = 0;
            if (passwordCharacters.Length > 0)
            {
                int charIndex = passwordCharacters.Length;

                while (charIndex-- > 0)
                {
                    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                    hash ^= passwordCharacters[charIndex];
                }
                // Main difference from spec, also hash with charcount
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters.Length;
                hash ^= (0x8000 | ('N' << 8) | 'K');
            }

            return Convert.ToString(hash, 16).ToUpperInvariant();
        }
    }
}
