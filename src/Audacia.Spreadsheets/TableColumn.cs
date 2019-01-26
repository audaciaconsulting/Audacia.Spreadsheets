using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets
{
    public class WorksheetTableColumn
    {
        public WorksheetTableColumn() { }

        public WorksheetTableColumn(string name) => Name = name;

        public static implicit operator WorksheetTableColumn(string name) => new WorksheetTableColumn(name);

        public string Name { get; set; }

        public bool IsIdColumn { get; set; }

        public CellFormatType Format { get; set; } = CellFormatType.Text;

        public CellBackgroundColourAttribute CellBackgroundFormat { get; set; }

        public CellTextColourAttribute CellTextFormat { get; set; }
    }
}
