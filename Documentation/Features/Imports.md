# Importing a spreadsheet
Step by step instructions to import data from a spreadsheet.

### Create an Importer class
```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;

public class BookImporter
{
    private IDictionary<string, number> _columnMap;
    
    public ICollection<Book> Import(Spreadsheet spreadsheet)
    {
        var books = new List<Book>();

        // Look at index or sheet name
        var sheet = spreadsheet.Worksheets[0];
        
        // Normally only one table
        var table = sheet.Tables.ElementAt(0);

        _columnMap = table.Columns.ToDictionary();
        
        foreach (var row in table.Rows)
        {
            var fields = row.GetFields();
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
    
    private string GetValue(string[] cellValues, string columnName)
    {
        if (!_columnMap.ContainsKey(columnName)) return string.Empty;
        
        var columnIndex = _columnMap[columnName];
        
        if (cellValues.Length <= columnIndex) return string.Empty;
        
        return cellValues[columnIndex];
    }
}
```

### Use the importer
```csharp
using System;
using System.IO;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;

var spreadsheet2 = default(Spreadsheet);
using (var fileStream = new FileStream(@".\Books.xlsx", FileMode.Open, FileAccess.Read))
{
    spreadsheet2 = Spreadsheet.FromStream(fileStream, includeHeaders: true, hasSubtotals: true);
    fileStream.Close();
}

var bookImporter = new BookImporter();
var importedBooks = bookImporter.Import(spreadsheet2);

```