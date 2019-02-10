## Migrating from previous libraries

Meaning Audacia.Spreadsheets.Export.
If you're working on a new project and are looking at this please leave now.

### Data Attributes

`using Audacia.Spreadsheets.Attributes;`

- CellBackgroundColour
- CellFormat
- CellTextColour
- IdColumn
- IgnoreDataMember
- HideHeader
- SubtotalHeader

### Creating a worksheet from a collection

```csharp
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;

var books = new [] { ... books ... };

var legacyWorksheet = books.ToWorksheet("Naughty Books");

var spreadsheet = Spreadsheet.FromWorksheets(legacyWorksheet);

var bytes = spreadsheet.Export();
```

