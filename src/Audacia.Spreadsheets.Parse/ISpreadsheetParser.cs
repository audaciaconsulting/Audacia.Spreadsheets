using System.IO;
using Audacia.Spreadsheets.Models;

namespace Audacia.Spreadsheets.Parse
{
    public interface ISpreadsheetParser
    {
        Spreadsheet GetSpreadsheetFromExcelFile(Stream stream, bool includeHeaders = true);
    }
}
