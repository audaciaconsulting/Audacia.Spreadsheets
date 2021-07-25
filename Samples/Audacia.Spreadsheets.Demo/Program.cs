using Audacia.Spreadsheets.Demo.Tasks;

namespace Audacia.Spreadsheets.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create custom worksheet exports & import by accessing the spreadsheet directly
            var customExample = new CustomExportImportTask();
            var bookFileBytes = customExample.Export();
            var books = customExample.Import(bookFileBytes);

            // Create a generic datasheet export and import
            var genericExample = new GenericExportImportTask();
            var accountFileBytes = genericExample.Export();
            var accounts = genericExample.Import(accountFileBytes);

            // Create a generic export with extended validation when importing
            var extendedExample = new ExtendedExportImportTask();
            var appointmentFileBytes = extendedExample.Export();
            var appointments = extendedExample.Import(appointmentFileBytes);

            // Create a generic export and import without column headers
            var noHeadersExample = new NoHeadersExportImportTask();
            var stockFileBytes = noHeadersExample.Export();
            var stock = noHeadersExample.Import(stockFileBytes);

            // Creating an export with multiple tables
        }
    }
}
