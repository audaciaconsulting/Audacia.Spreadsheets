using System.Collections.Generic;

namespace Audacia.Spreadsheets
{
    public class WorksheetTableRow
    {
        public int? Id { get; set; }
        public IList<TableCell> Cells { get; } = new List<TableCell>();
        
        public static WorksheetTableRow FromCells(IEnumerable<TableCell> cells, int? id)
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
