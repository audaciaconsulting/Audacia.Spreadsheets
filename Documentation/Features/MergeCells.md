# Merge cells

Cells can be merged by adding a cell range to the ``MergeCells`` list in the WorksheetBase class.

```csharp
var worksheet = new Worksheet
{
    SheetName = "My Worksheet"
};

worksheet.MergeCells.Add("A1:C1");
worksheet.MergeCells.Add("D2:C2");
```

```csharp
var worksheet = new MultiTableWorksheet
{
    SheetName = "Proposed Initial Allocation",
    MergeCells = ["A1:C1", "D2:C2"],
};
```
