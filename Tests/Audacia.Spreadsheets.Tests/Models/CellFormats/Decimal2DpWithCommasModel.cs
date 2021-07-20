using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class Decimal2DpWithCommasModel
    {
        [CellFormat(CellFormat.Decimal2DpWithCommas)]
        public decimal Value { get; set; }

        public static implicit operator Decimal2DpWithCommasModel(decimal d)
        {
            return new Decimal2DpWithCommasModel
            {
                Value = d
            };
        }
    }
}
