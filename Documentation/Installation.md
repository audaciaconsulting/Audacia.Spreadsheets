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

### Basic Exports:
For more info [see Exports](./Exports.md).
```csharp
var books = new [] { ... books ... };

// Create an exportable worksheet
var worksheet = new BookReport(books);

// Create a spreadsheet
var spreadsheet = Spreadsheet.FromWorksheets(worksheet1, worksheet2);

 // If you are a web project you can write to a byte array
var bytes = spreadsheet.Export();

// If you need to write to the file system you can pass it a stream
using (var fileStream = new FileStream(@".\Books.xlsx", FileMode.OpenOrCreate))
{
    spreadsheet.Write(fileStream);
    fileStream.Close();
}

```

### Basic Imports:
For more info [see Imports](./Imports.md).
```csharp
// Read from a stream.
var spreadsheet = default(Spreadsheet);
using (var fileStream = new FileStream(@".\Houses.xlsx", FileMode.Open, FileAccess.Read))
{
    spreadsheet = Spreadsheet.FromStream(fileStream);
    fileStream.Close();
}

// Create your own importer logic
var importer = new HouseImporter();
var houses = importer.Import(spreadsheet);

```