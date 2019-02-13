namespace Audacia.Spreadsheets.Validation
{
    /// <summary>
    /// An import error for when a cell value cannot be parsed.
    /// </summary>
    public class FieldParseError : RowImportError, IImportError
    {
        public string FieldName { get; }
        public string Value { get; }
        public string PossibleValues { get; }

        public FieldParseError(int rowNumber, string fieldName, string value, string possibleValues = null) 
            : base(rowNumber)
        {
            FieldName = fieldName;
            Value = value;
            PossibleValues = possibleValues;
        }

        public string GetMessage()
        {
            return $"{FieldName} is invalid{(!string.IsNullOrEmpty(PossibleValues) ? $", please use {PossibleValues}" : string.Empty)}.";
        }
    }
}