using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class IntegerModel
    {
        [CellFormat(CellFormat.Integer)]
        public int Value { get; set; }

        public static implicit operator IntegerModel(int i)
        {
            return new IntegerModel
            {
                Value = i
            };
        }
    }
}
