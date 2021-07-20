using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class BooleanYnModel
    {
        [CellFormat(CellFormat.BooleanYN)]
        public bool Value { get; set; }

        public static implicit operator BooleanYnModel(bool b)
        {
            return new BooleanYnModel
            {
                Value = b
            };
        }
    }
}
