using System.Collections.Generic;

namespace Audacia.Spreadsheets
{
    public class Worksheet
    {
        public int SheetIndex { get; set; }
        public string SheetName { get; set; }
        public IEnumerable<Table> Tables { get; set; }
        public WorksheetProtection WorksheetProtection { get; set; }
    }
}
