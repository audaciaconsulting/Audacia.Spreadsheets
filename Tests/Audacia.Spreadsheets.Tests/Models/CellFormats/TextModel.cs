using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Tests.Models.CellFormats
{
    public class TextModel
    {
        [CellFormat(CellFormat.Text)]
        public string Value { get; set; }

        public static implicit operator TextModel(string str)
        {
            return new TextModel
            {
                Value = str
            };
        }
    }
}
