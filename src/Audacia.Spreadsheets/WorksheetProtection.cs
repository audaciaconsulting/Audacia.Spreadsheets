using System.Collections.Generic;

namespace Audacia.Spreadsheets
{
    public class WorksheetProtection
    {
        public bool CanAddOrDeleteColumns { get; set; }
        public bool CanAddOrDeleteRows { get; set; }
        public string Password { get; set; }
    }
}
