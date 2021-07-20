using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class Decimal2DpFormat
    {
        [CellFormat(CellFormat.Decimal2Dp)]
        public decimal Value { get; set; }

        public static implicit operator Decimal2DpFormat(decimal d)
        {
            return new Decimal2DpFormat
            {
                Value = d
            };
        }
    }
}
