using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class IntegerWithCommasModel
    {
        [CellFormat(CellFormat.IntegerWithCommas)]
        public int Value { get; set; }

        public static implicit operator IntegerWithCommasModel(int i)
        {
            return new IntegerWithCommasModel
            {
                Value = i
            };
        }
    }
}
