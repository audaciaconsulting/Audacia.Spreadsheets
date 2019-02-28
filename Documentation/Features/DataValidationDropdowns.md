# Data Validation
You can add data validation dropdowns to your spreadsheets
## Static Data Validation
Simply create a StaticDropdown object and add it to your table's StaticDataValidations property:

```
spreadsheet.Worksheets.Add(new Worksheet()
{
    SheetName = "Test Sheet",

    StaticDataValidations = new List<StaticDropdown>()
    {
        new StaticDropdown()
        {
            AllowBlanks = true,
            Column = "A",
            Options = new List<string>(){ string.Join(",", listOfThings) }
        }
    },
}
```
In this example you can see that the list of items have been added as a comma-seperated list of items. This is because Excel wants to take in a comma-seperated list in it's data validation formula.

## Dependent Data Validation
Sometimes you will want to add a dropdown whose values can change based on a value in another cell. This is done using the DependentDataValidations property of the spreadsheet:

```
spreadsheet.Worksheets.Add(new Worksheet()
{
    SheetName = "Test Sheet",

    DependentDataValidations = new List<DependentDropdown>()
    {
        new DependentDropdown()
        {
            AllowBlanks = true,
            Column = "B",
            DependentColumn = "A"
        },
        new DependentDropdown()
        {
            AllowBlanks = true,
            Column = "C",
            DependentColumn = "B"
        },
        new DependentDropdown()
        {
            AllowBlanks = true,
            Column = "D",
            DependentColumn = "C"
        },
    }
}
```
To achieve this there are some critical things that need to be understood and set up:
1. The dropdown operates by converting the value of the dependent column to a Named Range and taking the values from whatever it finds in the spreasheet with that named range.
    - This is done using the INDIRECT() formula which converts a string value in to a named range and returns all cells in that named range. 
    - This means that you will need to create a Named Range for each value in the parent list which contains all of the values that you want to appear in the child list.
    - There is now also a `NamedRangeModel` which will help you to create a named range in the spreadsheet.
```
spreadsheet.NamedRanges.Add( 
    new NamedRangeModel()
    {
        Name = name.Replace(" ", "_"),
        SheetName = sheetName,
        StartCell = $"A{rowNumber}",
        EndCell = $"{rowLength.ToColumnLetter()}{rowNumber}"
    }
);
```

Once you have added the dropdowns and the named ranges, the spreadsheet should be able to Write itself in the correct order of operations.
