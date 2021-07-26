using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class ScientificModel
    {
        [CellFormat(CellFormat.Scientific)]
        public decimal Value { get; set; }

        public static implicit operator ScientificModel(decimal d)
        {
            return new ScientificModel
            {
                Value = d
            };
        }
    }
}
