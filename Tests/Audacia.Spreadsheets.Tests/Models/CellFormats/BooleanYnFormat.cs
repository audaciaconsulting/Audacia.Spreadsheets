using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class BooleanYnFormat
    {
        [CellFormat(CellFormat.BooleanYN)]
        public bool Value { get; set; }

        public static implicit operator BooleanYnFormat(bool b)
        {
            return new BooleanYnFormat
            {
                Value = b
            };
        }
    }
}
