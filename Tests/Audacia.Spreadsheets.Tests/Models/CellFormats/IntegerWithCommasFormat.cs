using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class IntegerWithCommasFormat
    {
        [CellFormat(CellFormat.IntegerWithCommas)]
        public int Value { get; set; }

        public static implicit operator IntegerWithCommasFormat(int i)
        {
            return new IntegerWithCommasFormat
            {
                Value = i
            };
        }
    }
}
