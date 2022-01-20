using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class FractionSmallModel
    {
        [CellFormat(CellFormat.FractionSmall)]
        public decimal Value { get; set; }

        public static implicit operator FractionSmallModel(decimal d)
        {
            return new FractionSmallModel
            {
                Value = d
            };
        }
    }
}
