# Excel Formulas
If you know a formula and want to put it in, you can.

For more information on creating worksheets see [Imports](/Imports.md).
```csharp
new TableCell("SUM(S:S)"),
{
    IsFormula = true
}
```

## Column Subtotals
If you've run the demo code from [Imports](/Imports.md), you probably noticed that there's an extra row above the column headers.
It is in fact a row dedicated to showing subtotals for specified columns.

If you choose to display a subtotal above a column, when exported the library will write a subtotal formula above the cell.
Please be aware that this will only sum numerical values.
```csharp
new TableColumn("Price (£)", CellFormat.Currency, displaySubtotal: true),
```

If you need to import a spreadsheet that has a subtotal row, this is also possible,
(this functionality may be subject to change).
```csharp
Spreadsheet.FromStream(fileStream, includeHeaders: true, hasSubtotals: true);
```