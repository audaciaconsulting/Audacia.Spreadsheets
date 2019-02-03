using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;
using Demo.Entities;

namespace Demo.Importers
{
    public class BookImporter
    {
        public bool Validate(Spreadsheet spreadsheet)
        {
            // The library doesn't contain any validation result objects
            // See Mark Dyer about import validation objects
            return true;
        }
        
        public ICollection<Book> Import(Spreadsheet spreadsheet)
        {
            var books = new List<Book>();

            // Look at index or sheet name
            var sheet = spreadsheet.Worksheets[0];
            
            // Normally only one table
            var table = sheet.Tables.ElementAt(0);

            // This won't work if the sheet you are importing has subtotals at the top
            var columns = table.Columns
                .Select((col, index) => (index, col.Name.Trim()))
                .ToArray();

            var rows = table.Rows
                .Select(row => row.Cells.Select(cell => cell.Value.ToString().Trim()).ToArray())
                .ToArray();

            // For this demo, I am using subtotals to test they are working
            var actualHeaders = rows.First();

            foreach (var item in rows.Skip(1))
            {
                // Have a column to property mapping dictionary
                // I'm being lazy because its a sunday afternoon right now
                var book = new Book
                {
                    Name = item[0],
                    Author = item[1],
                    Published = DateTime.ParseExact(item[2], "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                    Price = decimal.Parse(item[3]),
                    IsbnNumber = item[4]
                };

                books.Add(book);
            }

            return books;
        }

    }
}