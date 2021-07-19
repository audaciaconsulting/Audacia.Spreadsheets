using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class FractionSmallFormat
    {
        [CellFormat(CellFormat.FractionSmall)]
        public decimal Value { get; set; }
    }
}
