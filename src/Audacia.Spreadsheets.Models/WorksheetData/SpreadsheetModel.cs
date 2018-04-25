using System.Collections.Generic;

namespace Audacia.Spreadsheets.Models.WorksheetData
{
    public class SpreadsheetModel
    {
        public IEnumerable<WorksheetModel> Worksheets { get; set; } = new List<WorksheetModel>();
    }
}
