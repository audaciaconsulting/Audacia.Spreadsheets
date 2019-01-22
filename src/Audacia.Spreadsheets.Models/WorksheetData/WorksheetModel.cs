using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class WorksheetModel
    {
        public int SheetIndex { get; set; }
        public string SheetName { get; set; }
        public IEnumerable<TableModel> Tables { get; set; }
        public WorksheetProtection WorksheetProtection { get; set; }
    }
}
