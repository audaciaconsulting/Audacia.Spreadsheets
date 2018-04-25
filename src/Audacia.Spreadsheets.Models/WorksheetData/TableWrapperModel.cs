using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class TableWrapperModel
    {
        public IEnumerable<TableColumnModel> Columns { get; set; }
        public IEnumerable<TableRowModel> Rows { get; set; }
    }
}
