namespace Audacia.Spreadsheets
{
    /// <summary>
    /// Cell Data Type. This is a copy of
    /// <see cref="DocumentFormat.OpenXml.Spreadsheet.CellValues"/>
    /// but as constant strings rather than enums.
    /// </summary>
    public static class DataType
    {
        public const string Boolean = "b";
        
        public const string Date = "d";
        
        public const string Error = "e";
        
        public const string InlineString = "inlineStr";
        
        public const string Number = "n";
        
        public const string SharedString = "s";
        
        public const string String = "str";
    }
}
