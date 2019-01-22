using Audacia.Spreadsheets.Models.Enums;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class TableCellModel
    {
        public object Value { get; set; }
        public string FillColour { get; set; }
        public string TextColour { get; set; }
        public bool IsFormula { get; set; }
    }
}
