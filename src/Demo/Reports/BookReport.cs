using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;
using Demo.Entities;

namespace Demo.Reports
{
    public class BookReport : Worksheet
    {
        public BookReport() { }

        public BookReport(ICollection<Book> source)
        {
            var table = new Table(true);
            var rows = source.Select(FromBook);
            
            table.Columns.AddRange(Columns);
            table.Rows.AddRange(rows);
        }

        private static IEnumerable<TableColumn> Columns => new[]
        {
            new TableColumn("Name"),
            new TableColumn("Author"),
            new TableColumn("Published"),
            new TableColumn("Price (£)"),
            new TableColumn("ISBN Number")
        };
        
        private static TableRow FromBook(Book book)
        {
            return TableRow.FromCells(new []
            {
                new TableCell(book.Name),
                new TableCell(book.Author),
                new TableCell(book.Published),
                new TableCell(book.Price),
                new TableCell(book.IsbnNumber)
            }, null);
        }
    }
}