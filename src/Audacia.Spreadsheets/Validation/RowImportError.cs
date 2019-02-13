namespace Audacia.Spreadsheets.Validation
{
    /// <summary>
    /// An import error for a single row.
    /// </summary>
    public abstract class RowImportError
    {
        public int RowNumber { get; }

        public RowImportError(int rowNumber)
        {
            RowNumber = rowNumber;
        }
    }
}