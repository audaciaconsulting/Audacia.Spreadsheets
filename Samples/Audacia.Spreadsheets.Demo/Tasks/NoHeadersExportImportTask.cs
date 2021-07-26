using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using Audacia.Spreadsheets.Demo.Models;

namespace Audacia.Spreadsheets.Demo.Tasks
{
    /// <summary>
    /// Example of exporting and importing a dataset with generic logic and no column headers.
    /// </summary>
    public class NoHeadersExportImportTask
    {
        private Stock[] Dataset { get; } = new[]
        {
            new Stock { PartNumber = "PWE-376", Condition = "NEW", Manufacturer = "HP",     StockLevel = 7 },
            new Stock { PartNumber = "HDJ-213", Condition = "REF", Manufacturer = "Cisco",  StockLevel = 8 },
            new Stock { PartNumber = "OTN-676", Condition = "NEW", Manufacturer = "Lenovo", StockLevel = 3 },
            new Stock { PartNumber = "RMS-982", Condition = "REF", Manufacturer = "Dell",   StockLevel = 5 },
        };

        public byte[] Export()
        {
            Console.WriteLine("\r\nStock: Export() Started");

            // Print out data that will be exported
            Console.WriteLine("Printing dataset.");
            foreach (var a in Dataset)
            {
                Console.WriteLine(a);
            }

            // Create an exportable worksheet
            Console.WriteLine("Creating an exportable worksheet.");
            var worksheet = Dataset.ToWorksheet(includeHeaders: false);

            // Add all your worksheets into a spreadsheet
            Console.WriteLine("Creating a spreadsheet.");
            var spreadsheet = Spreadsheet.FromWorksheets(worksheet);

            // Export your spreadsheet
            // .Export() can write to a byte[] or a filepath, to write to a stream using .Write()
            Console.WriteLine("Exporting spreadsheet to a byte array.");
            var bytes = spreadsheet.Export();

            Console.WriteLine("Accounts: Export() Completed\r\n");
            return bytes;
        }

        public ICollection<Stock> Import(byte[] fileBytes)
        {
            Console.WriteLine("\r\nStock: Import() Started");

            // Read in spreadsheet, supports .FromBytes() and .FromStream()
            Console.WriteLine("Reading spreadsheet from bytes.");
            var spreadsheet = Spreadsheet.FromBytes(fileBytes, includeHeaders: false);

            // Parse first worksheet on the spreadsheet
            Console.WriteLine("Converting from spreadsheet back to collection of \"Stock\"");
            var stockImporter = new WorksheetImporter<Stock>()
                .SkipColumnHeaderMapping()
                .MapColumn("Part Number", x => x.PartNumber)
                .MapColumn("Manufacturer", x => x.Manufacturer)
                .MapColumn("Condition", x => x.Condition)
                .MapColumn("Stock Level", x => x.StockLevel);

            var importRows = stockImporter
                .ParseWorksheet(spreadsheet.Worksheets[0])
                .ToArray();

            // Print validation errors from failure to parse data
            if (importRows.Any(r => !r.IsValid))
            {
                Console.WriteLine("Import Failed: see below errors");
                foreach (var invalidRow in importRows.Where(r => !r.IsValid))
                {
                    Console.WriteLine("Row {0}:", invalidRow.RowId);

                    foreach (var fieldError in invalidRow.ImportErrors)
                    {
                        Console.WriteLine("\t{0}", fieldError.GetMessage());
                    }
                }
            }

            var stock = importRows
                .Where(r => r.IsValid)
                .Select(r => r.Data)
                .ToList();

            // Print out data that was imported
            Console.WriteLine("Printing dataset.");
            foreach (var a in stock)
            {
                Console.WriteLine(a);
            }

            Console.WriteLine("Stock: Import() Completed\r\n");
            return stock;
        }
    }
}
