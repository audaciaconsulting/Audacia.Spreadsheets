namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class BooleanModel
    {
        public bool Value { get; set; }

        public static implicit operator BooleanModel(bool b)
        {
            return new BooleanModel
            {
                Value = b
            };
        }
    }
}
