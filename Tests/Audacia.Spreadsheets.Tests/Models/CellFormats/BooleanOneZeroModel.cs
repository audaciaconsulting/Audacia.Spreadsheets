using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class BooleanOneZeroModel
    {
        [CellFormat(CellFormat.BooleanOneZero)]
        public bool Value { get; set; }

        public static implicit operator BooleanOneZeroModel(bool b)
        {
            return new BooleanOneZeroModel
            {
                Value = b
            };
        }
    }
}
