using Audacia.Spreadsheets.Demo.Tasks;

namespace Audacia.Spreadsheets.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Creating custom importers & exporters
            var customImporterTask = new CustomExportImportTask();
            var bookFileBytes = customImporterTask.Export();
            var books = customImporterTask.Import(bookFileBytes);

            // Creating a generic datasheet export
            var genericImporterTask = new GenericExportImportTask();
            var accountFileBytes = genericImporterTask.Export();
            var accounts = genericImporterTask.Import(accountFileBytes);

            // Creating a generic export with extended validation when importing
            var extendedImporterTask = new ExtendedExportImportTask();
            var appointmentFileBytes = extendedImporterTask.Export();
            var appointments = extendedImporterTask.Import(appointmentFileBytes);

            // Creating a generic export and import without column headers


            // Creating an export with multiple tables
        }
    }
}
