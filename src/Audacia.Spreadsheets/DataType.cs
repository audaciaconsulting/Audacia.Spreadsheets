namespace Audacia.Spreadsheets
{
    /// <summary>
    /// Cell Data Type. This is a copy of
    /// <see cref="DocumentFormat.OpenXml.Spreadsheet.CellValues"/>
    /// but using strings rather than enum values.
    /// </summary>
    public class DataType
    {
        private readonly string _value;

        private DataType() { }

        private DataType(string value) => _value = value;
        
        public static implicit operator string(DataType source) => source.ToString();

        public override string ToString() => _value;

        public static DataType Boolean { get; } = new DataType("b");

        public static DataType Date { get; } = new DataType("d");
        
        public static DataType Error { get; } = new DataType("e");
        
        public static DataType InlineString { get; } = new DataType("inlineStr");
        
        public static DataType Number { get; } = new DataType("n");
        
        public static DataType SharedString { get; } = new DataType("s");
        
        public static DataType String { get; } = new DataType("str");
    }
}
