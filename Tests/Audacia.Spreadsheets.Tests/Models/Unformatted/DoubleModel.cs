namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class DoubleModel
    {
        public double Value { get; set; }

        public static implicit operator DoubleModel(double d)
        {
            return new DoubleModel
            {
                Value = d
            };
        }
    }
}
