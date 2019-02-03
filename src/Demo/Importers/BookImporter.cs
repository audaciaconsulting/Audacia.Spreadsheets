using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;
using Demo.Entities;
using DocumentFormat.OpenXml.Spreadsheet;

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

            var columns = table.Columns
                .Select((col, index) => (index, col.Name.Trim()))
                .ToDictionary(c => c.Item2, c => c.Item1);

            var rows = table.Rows
                .Select(row => row.Cells.Select(cell => cell.Value.ToString().Trim()).ToArray())
                .ToArray();

            string GetValue(string[] cells, string columnName)
            {
                if (!columns.ContainsKey(columnName)) return string.Empty;
                
                var columnIndex = columns[columnName];
                
                if (cells.Length <= columnIndex) return string.Empty;
                
                return cells[columnIndex];
            }

            foreach (var item in rows)
            {
                var publishDate = GetValue(item, "Published");
                
                var book = new Book
                {
                    Name = GetValue(item, "Name"),
                    Author = GetValue(item, "Author"),
                    Published = DateTime.ParseExact(publishDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                    Price = decimal.Parse(GetValue(item, "Price (£)")),
                    IsbnNumber = GetValue(item, "ISBN Number")
                };

                books.Add(book);
            }

            return books;
        }

    }
}