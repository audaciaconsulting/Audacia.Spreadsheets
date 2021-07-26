using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class PercentageModel
    {
        [CellFormat(CellFormat.Percentage)]
        public decimal Value { get; set; }

        public static implicit operator PercentageModel(decimal d)
        {
            return new PercentageModel
            {
                Value = d
            };
        }
    }
}
