using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class Percentage2DpModel
    {
        [CellFormat(CellFormat.PercentageTwoDp)]
        public decimal Value { get; set; }

        public static implicit operator Percentage2DpModel(decimal d)
        {
            return new Percentage2DpModel
            {
                Value = d
            };
        }
    }
}
