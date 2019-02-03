using System;
using System.IO;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;
using Demo.Entities;
using Demo.Importers;
using Demo.Reports;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Creating dataset \"Books\".");
            var books = new[]
            {
                new Book { IsbnNumber = "086140324X", Author = "Terry Pratchet", Price = 9.99m, Published = DateTime.Now, Name = "The Colour of Magic" },
                new Book { IsbnNumber = "0861402030", Author = "Terry Pratchet", Price = 10m,   Published = DateTime.Now, Name = "The Light Fantastic" },
                new Book { IsbnNumber = "0304364258", Author = "Terry Pratchet", Price = 8.99m, Published = DateTime.Now, Name = "Equal Rites" },
                new Book { IsbnNumber = "0552152617", Author = "Terry Pratchet", Price = 20m,   Published = DateTime.Now, Name = "Mort" },
                new Book { IsbnNumber = "0575042176", Author = "Terry Pratchet", Price = 5m, Published = DateTime.Now, Name = "Sourcery" }
            };

            // Create an exportable worksheet
            Console.WriteLine("Creating an exportable worksheet.");
            var worksheet = new BookReport(books);
            
            // If you are migrating from a previous library
            // we still have .ToWorksheet() but it's now frowned upon
            Console.WriteLine("Creating an legacy worksheet.");
            var legacyWorksheet = books.ToWorksheet("Naughty Books");
            
            // Add all your worksheets into a spreadsheet
            Console.WriteLine("Creating a spreadsheet.");
            var spreadsheet = Spreadsheet.FromWorksheets(worksheet, legacyWorksheet);
            
            // If you are a web project or don't care you can write to a byte array
            Console.WriteLine("Exporting spreadsheet to a byte array.");
            var bytes = spreadsheet.Export();

            // If you need to write to the file system you can pass it a stream
            Console.WriteLine("Writing spreadsheet to a stream.");
            using (var fileStream = new FileStream(@".\Books.xlsx", FileMode.OpenOrCreate))
            {
                spreadsheet.Write(fileStream);
                fileStream.Close();
            }
            
            // If you want to parse a spreadsheet
            Console.WriteLine("Reading spreadsheet from a stream.");
            var spreadsheet2 = default(Spreadsheet);
            using (var fileStream = new FileStream(@".\Books.xlsx", FileMode.Open, FileAccess.Read))
            {
                spreadsheet2 = Spreadsheet.FromStream(fileStream, includeHeaders: true);
                fileStream.Close();
            }

            Console.WriteLine("Converting from spreadsheet back to collection of \"Books\"");
            var bookImporter = new BookImporter();
            if (bookImporter.Validate(spreadsheet2))
            {
                var importedBooks = bookImporter.Import(spreadsheet2);

                foreach (var b in importedBooks)
                {
                    Console.WriteLine(b);  
                }
            }
            
            Console.WriteLine("Press any key to close...");
            Console.ReadKey();
        }
    }
}
