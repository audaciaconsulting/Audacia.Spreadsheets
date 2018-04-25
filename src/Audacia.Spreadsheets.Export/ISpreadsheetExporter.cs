using Audacia.Spreadsheets.Models.WorksheetData;

namespace Audacia.Spreadsheets.Export
{
    public interface ISpreadsheetExporter
    {
        byte[] ExportSpreadsheetBytes(SpreadsheetModel model);
    }
}
