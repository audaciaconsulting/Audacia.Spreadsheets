# Audacia.Spreadsheets

Utilities for importing and generating spreadsheets.

## Pre-Installation

Make sure you have setup the Audacia VSTS nuget registry.
Either use the [adc tool](https://dev.azure.com/audacia/Audacia.DevOps/_git/Audacia.CommandLine?path=%2FREADME.md&version=GBmaster) or follow the instructions [here](https://docs.microsoft.com/en-gb/azure/devops/artifacts/nuget/consume?view=azure-devops&viewFallbackFrom=vsts&tabs=new-nav)

## Installation

If you are using the current template project you can skip this step.
Now that you have the registries setup you can install the package.

```powershell
# installation with nuget
Install-Package Audacia.Spreadsheets
```

### Basic Exports

For more info [see Exports](./Features/Exports.md), or the **Audacia.Spreadsheets.Demo** project.

```csharp
using Audacia.Spreadsheets;
using Audacia.Spreadsheets.Extensions;
...

var books = new Book[] { ... };

// Create an exportable worksheet
var worksheet = books.ToWorksheet();

// Add all your worksheets into a spreadsheet
var spreadsheet = Spreadsheet.FromWorksheets(worksheet);

// .Export() can write to a byte[] or a filepath, alternatively use .Write() to write to a stream
spreadsheet.Export("./books.xlsx");
```

### Basic Imports

For more info [see Imports](./Features/Imports.md), or the **Audacia.Spreadsheets.Demo** project.

```csharp
using System.Linq;
using Audacia.Spreadsheets;
...

// Alternatively you can read from a byte[] using .FromBytes() or stream using .FromStream()
var spreadsheet = Spreadsheet.FromFilePath("./books.xlsx");

// Inherit from WorksheetImporter<T> to implement your own custom parsing and/or validation logic per row
var importer = new WorksheetImporter<Book>();

var importedRows = importer.ParseWorksheet(spreadsheet.Worksheet[0]).ToArray();

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

### Migrating from previous libraries

Meaning Audacia.Spreadsheets.Export, please don't use this on newer projects.
Here are the names and namespaces for classes so your code can continue to work.

#### Legacy Data Attributes

`using Audacia.Spreadsheets.Attributes;`

- CellBackgroundColour
- CellFormat
- CellTextColour
- IdColumn
- IgnoreDataMember
- HideHeader
- SubtotalHeader
