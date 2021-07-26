# Importing a spreadsheet

Step by step instructions to import data from a spreadsheet.

## Create an Importer class

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Validation;

public class BookImporter : WorksheetImporter<Book>
{
    private readonly IDatabaseContext _dbContext;
    private readonly HashSet<int> _isbnChecksums = new HashSet<int>();

    public BookImporter(IDatabaseContext dbContext)
    {
        _dbContext = dbContext;
    }

    protected override IEnumerable<IImportError> ValidateRow(Book model)
    {
        var importErrors = new List<IImportError>();

        if (string.IsNullOrWhiteSpace(model.Name))
        {
            importErrors.Add(new FieldValidationError(GetRowNumber(), new[]
            {
                new ValidationResult(GetColumnHeader(x => x.Name), "This field is required.")
            }));
        }

        var checksum = row.IsbnNumber.GetHashCode();
        if (!_isbnChecksums.Add(checksum))
        {
            importErrors.Add(new DuplicateKeyError(GetRowNumber(), GetColumnHeader(x => x.IsbnNumber), row.IsbnNumber));
        }
        else if (_dbContext.Books.Any(b => b.IsbnNumber == model.IsbnNumber)) 
        {
            importErrors.Add(new RecordExistsError(GetRowNumber(), GetColumnHeader(x => x.IsbnNumber), row.IsbnNumber));
        }

        return importErrors;
    }
}
```

## Use the importer

```csharp
using System.Linq;
using Audacia.Spreadsheets;
...

var spreadsheet = Spreadsheet.FromFilePath("./books.xlsx");

var worksheet = spreadsheet.Worksheets.FirstOrDefault(w => w.SheetName == "Books");

var importedRows = new BookImporter(dbContext)
    .ParseWorksheet(worksheet)
    .ToArray();

// Handle rows that failed to map to an object...
if (importedRows.Any(x => !x.IsValid))
{
    var invalidRows = importedRows
        .Where(x => !x.IsValid);
    ...

    return;
}

// Get parsed data from imported rows
var books = importedRows
    .Where(b => b.IsValid)
    .Select(b => b.Data)
    .ToArray();
```
