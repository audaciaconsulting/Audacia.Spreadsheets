using System.Collections.Generic;
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

        public void Write(UInt32Value styleIndex, string reference, string dataType, OpenXmlWriter writer)
        {
            WriteCell(writer, styleIndex, reference, dataType, Value.ToString(), IsFormula);
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
