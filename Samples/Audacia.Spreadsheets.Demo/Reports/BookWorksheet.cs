using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Demo.Models;

namespace Audacia.Spreadsheets.Demo.Reports
{
    public class BookWorksheet : Worksheet
    {
        public BookWorksheet(ICollection<Book> source)
        {
            SheetName = "Books";
            ShowGridLines = true;
            HasAutofilter = true;
            Table = new Table(includeHeaders: true)
            {
                Columns = Columns.ToList(),
                Rows = source.Select(FromBook)
            };
        }

        private static IEnumerable<TableColumn> Columns => new[]
        {
            new TableColumn("Name", hasBorders: false),
            new TableColumn("Author", hasBorders: false),
            new TableColumn("Published", CellFormat.Date, hasBorders: false),
            new TableColumn("Price (£)", CellFormat.AccountingGBP, displaySubtotal: true, hasBorders: false),
            new TableColumn("ISBN Number", hasBorders: false),
            new TableColumn("Full", (CellFormat)30U, hasBorders: false),
            new TableColumn("Hours", CellFormat.TimeSpanHours, hasBorders: false),
            new TableColumn("Mins", CellFormat.TimeSpanMinutes, hasBorders: false)
        };
        
        private static TableRow FromBook(Book book)
        {
            return TableRow.FromCells(new []
            {
                new TableCell(book.Name, hasBorders: false),
                new TableCell(book.Author, hasBorders: false),
                new TableCell(book.Published, hasBorders: false),
                new TableCell(book.Price, hasBorders: false),
                new TableCell(book.IsbnNumber, hasBorders: false),
                new TableCell(book.Timespan, hasBorders: false),
                new TableCell(book.Timespan, hasBorders: false),
                new TableCell(book.Timespan, hasBorders: false)
            }, null);
        }
    }
}