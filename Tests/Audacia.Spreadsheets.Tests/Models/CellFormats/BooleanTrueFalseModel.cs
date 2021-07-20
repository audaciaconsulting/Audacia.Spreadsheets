using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class BooleanTrueFalseModel
    {
        [CellFormat(CellFormat.BooleanTrueFalse)]
        public bool Value { get; set; }

        public static implicit operator BooleanTrueFalseModel(bool b)
        {
            return new BooleanTrueFalseModel
            {
                Value = b
            };
        }
    }
}
