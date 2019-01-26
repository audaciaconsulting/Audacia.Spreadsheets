using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models
{
    public class Worksheet
    {
        public int SheetIndex { get; set; }
        public string SheetName { get; set; }
        public IEnumerable<WorksheetTable> Tables { get; set; }
        public WorksheetProtection WorksheetProtection { get; set; }
    }
}
