using System;
using System.Collections.Generic;
using System.Globalization;
using Audacia.Spreadsheets.Demo.Models;
using Audacia.Spreadsheets.Extensions;

namespace Audacia.Spreadsheets.Demo.Importers
{
    /// <summary>
    /// Example barebones importer, does not contain validation.
    /// </summary>
    public class BookImporter
    {
        public ICollection<Book> Import(Spreadsheet spreadsheet)
        {
            // Get the table from the first worksheet
            // When importing a spreadsheet there will only be one table per worksheet.
            var table = spreadsheet.Worksheets[0].GetTable();

            // Get a column map of column header to column index
            var columns = table.Columns.ToDictionary();

            var books = new List<Book>();
            foreach (var row in table.Rows)
            {
                // Get cell data
                var fields = row.GetFields();

                // Validate the row here

                // Create model object
                var publishDate = GetValue(fields, "Published");
                var book = new Book
                {
                    Name = GetValue(fields, "Name"),
                    Author = GetValue(fields, "Author"),
                    Published = DateTime.ParseExact(publishDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                    Price = decimal.Parse(GetValue(fields, "Price (£)")),
                    IsbnNumber = GetValue(fields, "ISBN Number")
                };

                books.Add(book);
            }

            return books;

            string GetValue(string[] fields, string columnName)
            {
                if (!columns.ContainsKey(columnName))
                {
                    return string.Empty;
                }
                
                var columnIndex = columns[columnName];

                if (fields.Length <= columnIndex)
                {
                    return string.Empty;
                }
                
                return fields[columnIndex];
            }
        }
    }
}