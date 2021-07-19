using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class IntegerWithCommasFormat
    {
        [CellFormat(CellFormat.IntegerWithCommas)]
        public int Value { get; set; }
    }
}
