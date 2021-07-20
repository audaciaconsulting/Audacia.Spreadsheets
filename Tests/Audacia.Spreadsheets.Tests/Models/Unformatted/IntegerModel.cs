namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class IntegerModel
    {
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
