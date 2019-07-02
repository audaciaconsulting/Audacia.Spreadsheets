using System;
using System.Collections.Generic;
using System.Globalization;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class TableCell
    {
        public TableCell() { }

        public TableCell(object value)
        {
            Value = value;
        }
        
        public TableCell(object value, bool isFormula)
        {
            Value = value;
            IsFormula = isFormula;
        }
        
        public TableCell(object value = null, bool isFormula = false, bool hasBorders = true, bool isBold = false)
        {
            Value = value;
            IsFormula = isFormula;
            HasBorders = hasBorders;
            IsBold = isBold;
        }

        public object Value { get; set; }

        public string FillColour { get; set; }

        public string TextColour { get; set; }

        public bool IsFormula { get; set; }

        public bool HasBorderTop { get; set; } = true;
        public bool HasBorderRight { get; set; } = true;
        public bool HasBorderBottom { get; set; } = true;
        public bool HasBorderLeft { get; set; } = true;

        // making still usable despite separating the border into 4 sections as may be required
        public bool HasBorders
        {
            get => HasBorderTop && HasBorderRight && HasBorderBottom && HasBorderLeft;
            set 
            { 
                HasBorderTop = value;
                HasBorderRight = value;
                HasBorderBottom = value;
                HasBorderLeft = value;
            }
        }

        public bool IsBold { get; set; }

        public CellStyle CellStyle(TableColumn column)
        {
            return new CellStyle
            {
                TextColour = IsBold ? 1U : 0U,
                BackgroundColour = 0U,
                BorderBottom = HasBorderBottom,
                BorderTop = HasBorderTop,
                BorderLeft = HasBorderLeft,
                BorderRight = HasBorderRight,
                Format = column.Format,
                HasWordWrap = Value is string && !IsFormula,
                IsEditable = IsEditable
            };
        }

        public bool IsEditable { get; set; }

        public void Write(UInt32Value styleIndex, CellFormat format, string reference, OpenXmlWriter writer)
        {
            (DataType dataType, string value) = GetDataTypeAndFormattedValue(format);
            WriteCell(writer, styleIndex, reference, dataType, value, IsFormula);
        }

        public static void WriteCell(OpenXmlWriter writer, UInt32Value styleIndex,
            string reference, DataType dataType, string value, bool isFormula)
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

        private Tuple<DataType, string> GetDataTypeAndFormattedValue(CellFormat format)
        {
            switch (Value)
            {
                case DateTime date:
                    return new Tuple<DataType, string>(DataType.Number, FormatDateTimeAsString(date));
                case DateTimeOffset date:
                    return new Tuple<DataType, string>(DataType.Number, FormatDateTimeAsString(date.LocalDateTime));
                case TimeSpan t:
                {
                    if (format == CellFormat.Text)
                    {
                        return new Tuple<DataType, string>(DataType.String, Value.ToString());
                    }
                    // Use the provided number format
                    return new Tuple<DataType, string>(DataType.Number, t.ToOADatePrecise().ToString(CultureInfo.CurrentCulture));
                } 
                case decimal dec:
                    return new Tuple<DataType, string>(DataType.Number, dec.ToString(CultureInfo.CurrentCulture));
                case double d:
                    return new Tuple<DataType, string>(DataType.Number, d.ToString(CultureInfo.CurrentCulture));
                case float f:
                    return new Tuple<DataType, string>(DataType.Number, f.ToString(CultureInfo.CurrentCulture));
                case int i:
                    return new Tuple<DataType, string>(DataType.Number, i.ToString(CultureInfo.CurrentCulture));
                case bool b:
                    return new Tuple<DataType, string>(DataType.String, FormatBooleanAsString(format, b));
                default:
                    return new Tuple<DataType, string>(DataType.String, Value?.ToString() ?? string.Empty);
            }
        }

        private static string FormatBooleanAsString(CellFormat format, bool value)
        {
            switch (format)
            {
                case CellFormat.BooleanYN: return value ? "Y" : "N";
                case CellFormat.BooleanYesNo: return value ? "Yes" : "No";
                case CellFormat.BooleanOneZero: return value ? "1" : "0";
                case CellFormat.BooleanTrueFalse: return value ? "True" : "False";
                default: return value.ToString();
            }
        }

        private static string FormatDateTimeAsString(DateTime value)
        {
            return !value.Equals(DateTime.MinValue)
                ? value.ToOADatePrecise().ToString(CultureInfo.CurrentCulture)
                : string.Empty;
        }
    }
}
