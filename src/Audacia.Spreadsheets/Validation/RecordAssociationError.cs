using System.Collections.Generic;

namespace Audacia.Spreadsheets.Validation
{
    /// <summary>
    /// To be used when the referenced record cannot be associated.
    /// </summary>
    public class RecordAssociationError : RowGroupImportError, IImportError
    {
        private string Type1 { get; }
        private string Type2 { get; }
        private string Value1 { get; }
        private string Value2 { get; }

        public RecordAssociationError(string type1, string type2, string value1, string value2, IEnumerable<int> rowNumbers) 
            : base(rowNumbers)
        {
            Type1 = type1;
            Type2 = type2;
            Value1 = value1;
            Value2 = value2;
        }

        public string GetMessage()
        {
            return $"{Type1} of {Value1} is not associated with {Type2} of {Value2}.";
        }
    }
}