using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class Decimal2DpModel
    {
        [CellFormat(CellFormat.Decimal2Dp)]
        public decimal Value { get; set; }

        public static implicit operator Decimal2DpModel(decimal d)
        {
            return new Decimal2DpModel
            {
                Value = d
            };
        }
    }
}
