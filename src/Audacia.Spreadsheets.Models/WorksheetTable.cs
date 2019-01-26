using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models
{
    public class WorksheetTable
    {
        public string StartingCellRef { get; set; }
        
        public WorksheetHeaderStyle HeaderStyle { get; set; }

        public bool IncludeHeaders { get; set; }

        // TODO JP: move logic for freeze rows into library
        /// <summary>Number of rows to freeze starting from the top row</summary>
        public int FreezeTopRows { get; set; }
        
        public IList<WorksheetTableColumn> Columns { get; } = new List<WorksheetTableColumn>();

        public IList<WorksheetTableRow> Rows { get; } = new List<WorksheetTableRow>();
    }
}
