# Exporting a spreadsheet
Step by step instructions to create and export a worksheet.

### Have an entity to export
```csharp
using System;
using Audacia.Spreadsheets;

public class Book
{
    public string Name { get; set; }
    public string Author { get; set; }
    public string IsbnNumber { get; set; }
    public DateTime Published { get; set; }
    public decimal Price { get; set; }

    public override string ToString()
    {
        return $"{Name}, {Author}, {Published:d}, {Price:C}, {IsbnNumber}";
    }
}
```

### Define a worksheet model
```csharp
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;

///<summary>Takes a collection of books and creates an exportable worksheet.</summary>
public class BookReport : Worksheet
{
    public BookReport(ICollection<Book> source)
    {
        SheetName = "Good Books";
        var table = new Table(includeHeaders: true);
        var rows = source.Select(FromBook);
        
        table.Columns.AddRange(Columns);
        table.Rows.AddRange(rows);
        Table = table;
    }

    ///<summary>Define columns and how they will be formatted.</summary>
    private static IEnumerable<TableColumn> Columns => new[]
    {
        new TableColumn("Name"),
        new TableColumn("Author"),
        new TableColumn("Published", CellFormat.Date),
        new TableColumn("Price (£)", CellFormat.Currency, true),
        new TableColumn("ISBN Number")
    };
    
    ///<summary>Converts a book to a row.</summary>
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
```

### Populate the worksheet and export it
```csharp
using System;
using System.IO;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;

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

    Console.WriteLine("Creating a worksheet.");
    var worksheet = new BookReport(books);
    
    Console.WriteLine("Creating a spreadsheet.");
    var spreadsheet = Spreadsheet.FromWorksheets(worksheet);
    
    // If you need to write to the file system you can pass it a stream
    Console.WriteLine("Writing spreadsheet to a stream.");
    using (var fileStream = new FileStream(@".\Books.xlsx", FileMode.OpenOrCreate))
    {
        spreadsheet.Write(fileStream);
        fileStream.Close();
    }
}
```
