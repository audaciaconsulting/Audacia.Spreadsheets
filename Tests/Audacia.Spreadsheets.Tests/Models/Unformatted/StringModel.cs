namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class StringModel
    {
        public string Value { get; set; }

        public static implicit operator StringModel(string str)
        {
            return new StringModel
            {
                Value = str
            };
        }
    }
}
