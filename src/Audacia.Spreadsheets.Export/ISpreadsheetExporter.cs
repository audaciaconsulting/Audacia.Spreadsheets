using Audacia.Spreadsheets.Models;

namespace Audacia.Spreadsheets.Export
{
    public interface ISpreadsheetExporter
    {
        byte[] ExportSpreadsheetBytes(Spreadsheet model);
    }
}
