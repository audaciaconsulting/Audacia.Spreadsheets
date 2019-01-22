using System;
using System.Collections.Generic;
using System.Text;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class WorksheetProtection
    {
        public bool CanAddOrDeleteColumns { get; set; }
        public bool CanAddOrDeleteRows { get; set; }
        public string Password { get; set; }
        public IEnumerable<string> EditableCellRanges { get; set; }
    }
}
