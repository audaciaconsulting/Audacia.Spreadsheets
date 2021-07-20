using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class FractionSmallFormat
    {
        [CellFormat(CellFormat.FractionSmall)]
        public decimal Value { get; set; }

        public static implicit operator FractionSmallFormat(decimal d)
        {
            return new FractionSmallFormat
            {
                Value = d
            };
        }
    }
}
