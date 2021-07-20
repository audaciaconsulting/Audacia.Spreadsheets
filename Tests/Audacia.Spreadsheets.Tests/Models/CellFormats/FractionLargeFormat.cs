using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class FractionLargeFormat
    {
        [CellFormat(CellFormat.FractionLarge)]
        public decimal Value { get; set; }

        public static implicit operator FractionLargeFormat(decimal d)
        {
            return new FractionLargeFormat
            {
                Value = d
            };
        }
    }
}
