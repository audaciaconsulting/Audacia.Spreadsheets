using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class Percentage2DpFormat
    {
        [CellFormat(CellFormat.Percentage2Dp)]
        public decimal Value { get; set; }

        public static implicit operator Percentage2DpFormat(decimal d)
        {
            return new Percentage2DpFormat
            {
                Value = d
            };
        }
    }
}
