namespace Audacia.Spreadsheets.Tests.Models.Unformatted
{
    public class EnumModel
    {
        public enum Shape
        {
            Circle,
            Triangle,
            Square,
            Pentagon,
            Hexagon
        }

        public Shape Value { get; set; }

        public static implicit operator EnumModel(Shape t)
        {
            return new EnumModel
            {
                Value = t
            };
        }
    }
}
