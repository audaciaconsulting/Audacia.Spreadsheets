using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class PercentageFormat2Dp
    {
        [CellFormat(CellFormat.Percentage2Dp)]
        public decimal Value { get; set; }
    }
}
