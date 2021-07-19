using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class Decimal2DpWithCommasFormat
    {
        [CellFormat(CellFormat.Decimal2DpWithCommas)]
        public decimal Value { get; set; }
    }
}
