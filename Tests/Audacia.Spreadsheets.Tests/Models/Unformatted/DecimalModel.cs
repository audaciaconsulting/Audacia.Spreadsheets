namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class DecimalModel
    {
        public decimal Value { get; set; }

        public static implicit operator DecimalModel(decimal d)
        {
            return new DecimalModel
            {
                Value = d
            };
        }
    }
}
