namespace Audacia.Spreadsheets
{
    public class WorksheetTableCell
    {   
        public WorksheetTableCell() { }

        public WorksheetTableCell(object value) => Value = value;

        public object Value { get; set; }

        public string FillColour { get; set; }

        public string TextColour { get; set; }
        
        public bool IsFormula { get; set; }
    }
}
