using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class BooleanOneZeroFormat
    {
        [CellFormat(CellFormat.BooleanOneZero)]
        public bool Value { get; set; }

        public static implicit operator BooleanOneZeroFormat(bool b)
        {
            return new BooleanOneZeroFormat
            {
                Value = b
            };
        }
    }
}
