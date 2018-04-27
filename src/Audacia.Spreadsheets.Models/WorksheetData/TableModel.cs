namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class TableModel
    {
        public string StartingCellRef { get; set; }
        public bool IncludeHeaders { get; set; }
        public TableWrapperModel Data { get; set; }
        public SpreadsheetHeaderStyle HeaderStyle { get; set; }
    }
}
