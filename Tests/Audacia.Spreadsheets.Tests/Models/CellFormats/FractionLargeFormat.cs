using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class FractionLargeFormat
    {
        [CellFormat(CellFormat.FractionLarge)]
        public decimal Value { get; set; }
    }
}
