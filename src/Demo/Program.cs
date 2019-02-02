using System;
using System.IO;
using Audacia.Spreadsheets;
using Demo.Entities;
using Demo.Reports;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var books = new[]
            {
                new Book { IsbnNumber = "086140324X", Author = "Terry Pratchet", Price = 9.99m, Published = DateTime.Now, Name = "The Colour of Magic" },
                new Book { IsbnNumber = "0861402030", Author = "Terry Pratchet", Price = 10m,   Published = DateTime.Now, Name = "The Light Fantastic" },
                new Book { IsbnNumber = "0304364258", Author = "Terry Pratchet", Price = 8.99m, Published = DateTime.Now, Name = "Equal Rites" },
                new Book { IsbnNumber = "0552152617", Author = "Terry Pratchet", Price = 20m,   Published = DateTime.Now, Name = "Mort" },
                new Book { IsbnNumber = "0575042176", Author = "Terry Pratchet", Price = 5m, Published = DateTime.Now, Name = "Sourcery" }
            };

            // Create an exportable worksheet
            var worksheet = new BookReport(books);
            
            // Add all your worksheets into a spreadsheet
            var spreadsheet = Spreadsheet.FromWorksheets(worksheet);
            
            // Write to byte array
            var bytes = spreadsheet.Export();

            // Or write to file
            using (var fileStream = new FileStream(@".\Books.xlsx", FileMode.OpenOrCreate))
            {
                spreadsheet.Write(fileStream);
                fileStream.Close();
            }
        }
    }
}
