using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class BooleanTrueFalseFormat
    {
        [CellFormat(CellFormat.BooleanTrueFalse)]
        public bool Value { get; set; }

        public static implicit operator BooleanTrueFalseFormat(bool b)
        {
            return new BooleanTrueFalseFormat
            {
                Value = b
            };
        }
    }
}
