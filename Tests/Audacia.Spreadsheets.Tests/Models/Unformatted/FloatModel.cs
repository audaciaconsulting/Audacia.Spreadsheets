namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class FloatModel
    {
        public float Value { get; set; }

        public static implicit operator FloatModel(float f)
        {
            return new FloatModel
            {
                Value = f
            };
        }
    }
}
