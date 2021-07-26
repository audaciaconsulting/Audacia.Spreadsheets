using System;
using System.Collections.Generic;
using Audacia.Spreadsheets.Demo.Importers;
using Audacia.Spreadsheets.Demo.Models;
using Audacia.Spreadsheets.Demo.Reports;

namespace Audacia.Spreadsheets.Demo.Tasks
{
    /// <summary>
    /// Example of exporting and importing a dataset with customisable logic.
    /// </summary>
    public class CustomExportImportTask
    {
        private Book[] Dataset { get; } = new[]
        {
            new Book { IsbnNumber = "086140324X", Author = "Terry Pratchet", Published = DateTime.Now, Price = 9.99m, Name = "The Colour of Magic", },
            new Book { IsbnNumber = "0861402030", Author = "Terry Pratchet", Published = DateTime.Now, Price = 10m,   Name = "The Light Fantastic", },
            new Book { IsbnNumber = "0304364258", Author = "Terry Pratchet", Published = DateTime.Now, Price = 8.99m, Name = "Equal Rites",         },
            new Book { IsbnNumber = "0552152617", Author = "Terry Pratchet", Published = DateTime.Now, Price = 20m,   Name = "Mort",                },
            new Book { IsbnNumber = "0575042176", Author = "Terry Pratchet", Published = DateTime.Now, Price = 5m,    Name = "Sourcery",            }
        };

        public byte[] Export()
        {
            Console.WriteLine("\r\nBooks: Export() Started");

            // Print out data that will be exported
            Console.WriteLine("Printing dataset.");
            foreach (var b in Dataset)
            {
                Console.WriteLine(b);
            }

            // Create an exportable worksheet
            Console.WriteLine("Creating an exportable worksheet.");
            var worksheet = new BookWorksheet(Dataset);

            // Add all your worksheets into a spreadsheet
            Console.WriteLine("Creating a spreadsheet.");
            var spreadsheet = Spreadsheet.FromWorksheets(worksheet);

            // Export your spreadsheet
            // .Export() can write to a byte[] or a filepath, to write to a stream using .Write()
            Console.WriteLine("Exporting spreadsheet to a byte array.");
            var bytes = spreadsheet.Export();

            Console.WriteLine("Books: Export() Completed");
            return bytes;
        }

        public ICollection<Book> Import(byte[] fileBytes)
        {
            Console.WriteLine("\r\nBooks: Import() Started");

            // Read in spreadsheet, supports .FromBytes() and .FromStream()
            Console.WriteLine("Reading spreadsheet from bytes.");
            var spreadsheet = Spreadsheet.FromBytes(fileBytes, hasSubtotals: true);

            // Parse first worksheet on the spreadsheet
            Console.WriteLine("Converting from spreadsheet back to collection of \"Books\"");
            var bookImporter = new BookImporter();
            var books = bookImporter.Import(spreadsheet);

            // Print out data that was imported
            Console.WriteLine("Printing dataset.");
            foreach (var b in books)
            {
                Console.WriteLine(b);
            }

            Console.WriteLine("Books: Import() Completed\r\n");
            return books;
        }
    }
}