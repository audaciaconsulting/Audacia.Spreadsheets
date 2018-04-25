using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class TableRowModel
    {
        public int? Id { get; set; }
        public IEnumerable<TableCellModel> Cells { get; set; }
    }
}
