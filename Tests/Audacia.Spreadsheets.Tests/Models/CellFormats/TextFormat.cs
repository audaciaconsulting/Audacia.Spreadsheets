using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TextFormat
    {
        [CellFormat(CellFormat.Text)]
        public string Value { get; set; }

        public static implicit operator TextFormat(string str)
        {
            return new TextFormat
            {
                Value = str
            };
        }
    }
}
