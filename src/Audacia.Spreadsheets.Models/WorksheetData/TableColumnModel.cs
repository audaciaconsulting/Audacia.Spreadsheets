using Audacia.Spreadsheets.Models.Attributes;
using Audacia.Spreadsheets.Models.Enums;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class TableColumnModel
    {
        public string Name { get; set; }
        public bool IsIdColumn { get; set; }
        public CellFormatType Format { get; set; } = CellFormatType.Text;
        public CellBackgroundColourAttribute CellBackgroundFormat { get; set; }
        public CellTextColourAttribute CellTextFormat { get; set; }
    }
}
