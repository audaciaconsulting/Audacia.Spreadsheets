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

        public TableCell(object value) => Value = value;

        public object Value { get; set; }

        public string FillColour { get; set; }

        public string TextColour { get; set; }
        
        public bool IsFormula { get; set; }

        private Tuple<string, string> GetDataTypeAndFormattedValue()
        {
            switch (Value)
            {
                case DateTime date:
                {
                    var dateString = string.Empty;
                    if (!date.Equals(DateTime.MinValue))
                    {
                        dateString = date.ToOADatePrecise().ToString(new CultureInfo("en-US"));
                    }

                    return new Tuple<string, string>(DataType.Date, dateString);
                }
                case decimal dec:
                    return new Tuple<string, string>(DataType.Numeric, dec.ToString(CultureInfo.CurrentCulture));
                case double d:
                    return new Tuple<string, string>(DataType.Numeric, d.ToString(CultureInfo.CurrentCulture));
                case int i:
                    return new Tuple<string, string>(DataType.Numeric, i.ToString(CultureInfo.CurrentCulture));
                default:
                    return new Tuple<string, string>(DataType.String, Value.ToString());
            }
        }
        
        public void Write(UInt32Value styleIndex, string reference, OpenXmlWriter writer)
        {
            (string dataType, string value) = GetDataTypeAndFormattedValue();
            WriteCell(writer, styleIndex, reference, dataType, value, IsFormula);
        }

        public static void WriteCell(OpenXmlWriter writer, UInt32Value styleIndex,
            string reference, string dataType, string value, bool isFormula)
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
    }
}
