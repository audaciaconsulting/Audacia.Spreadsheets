using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class ScientificFormat
    {
        [CellFormat(CellFormat.Scientific)]
        public decimal Value { get; set; }
    }
}
