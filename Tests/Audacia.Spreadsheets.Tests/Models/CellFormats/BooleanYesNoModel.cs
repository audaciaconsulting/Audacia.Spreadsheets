using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class BooleanYesNoModel
    {
        [CellFormat(CellFormat.BooleanYesNo)]
        public bool Value { get; set; }

        public static implicit operator BooleanYesNoModel(bool b)
        {
            return new BooleanYesNoModel
            {
                Value = b
            };
        }
    }
}
