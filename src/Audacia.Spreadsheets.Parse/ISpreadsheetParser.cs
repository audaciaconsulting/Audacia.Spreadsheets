using Audacia.Spreadsheets.Models.WorksheetData;
using System.IO;

namespace Audacia.Spreadsheets.Parse
{
    public interface ISpreadsheetParser
    {
        SpreadsheetModel GetSpreadsheetFromExcelFile(Stream stream, bool includeHeaders = true);
    }
}
