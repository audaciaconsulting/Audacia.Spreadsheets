using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class IntegerFormat
    {
        [CellFormat(CellFormat.Integer)]
        public int Value { get; set; }
    }
}
