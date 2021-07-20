using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class ScientificFormat
    {
        [CellFormat(CellFormat.Scientific)]
        public decimal Value { get; set; }

        public static implicit operator ScientificFormat(decimal d)
        {
            return new ScientificFormat
            {
                Value = d
            };
        }
    }
}
