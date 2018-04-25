using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class TableWrapperModel
    {
        public string Area { get; set; }
        public string Controller { get; set; }
        public string Action { get; set; }
        public IEnumerable<TableColumnModel> Columns { get; set; }
        public IEnumerable<TableRowModel> Rows { get; set; }
    }
}
