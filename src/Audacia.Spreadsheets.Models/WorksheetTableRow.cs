using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models
{
    public class WorksheetTableRow
    {
        public int? Id { get; set; }
        public IList<WorksheetTableCell> Cells { get; } = new List<WorksheetTableCell>();
        
        public static WorksheetTableRow FromCells(IEnumerable<WorksheetTableCell> cells, int? id)
        {
            var row = new WorksheetTableRow
            {
                Id = id
            };
            
            foreach (var cell in cells)
            {
                row.Cells.Add(cell);
            }
            
            return row;
        }
    }
}
