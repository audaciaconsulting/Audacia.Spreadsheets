using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class FractionLargeModel
    {
        [CellFormat(CellFormat.FractionLarge)]
        public decimal Value { get; set; }

        public static implicit operator FractionLargeModel(decimal d)
        {
            return new FractionLargeModel
            {
                Value = d
            };
        }
    }
}
