using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class BooleanYesNoFormat
    {
        [CellFormat(CellFormat.BooleanYesNo)]
        public bool Value { get; set; }

        public static implicit operator BooleanYesNoFormat(bool b)
        {
            return new BooleanYesNoFormat
            {
                Value = b
            };
        }
    }
}
