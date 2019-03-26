# Cell Borders

Cell borders are enabled by default.
Borders may be disabled on a cell by cell basis.

If a cell has a border then it will surround the entire cell.
There is currently no ability to configure individual borders on a cell. 

```csharp
using Audacia.Spreadsheets;

var column1 = new TableColumn(hasBorders: false);

var column2 = new TableColumn()
{
    HasBorders = false
};

var cell1 = new TableCell(hasBorders: false);

var cell2 = new TableCell()
{
    HasBorders = false
};
```

# Gridlines

Gridlines may be enabled or disabled on worksheets.
Gridlines are disabled by default.

```csharp
using Audacia.Spreadsheets;

class MyWorksheet : Worksheet
{
    public MyWorksheet() 
    {
        ShowGridlines = true;
    }
}
```
