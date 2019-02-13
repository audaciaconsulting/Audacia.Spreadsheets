# Locking a worksheet

You can add WorksheetProtection to worksheets.

If done right this will password protect a region of cells or an entire worksheet.

```csharp
public class HouseReport : Worksheet
{
    public HouseReport(...)
    {
        ...
        
        WorksheetProtection = new WorksheetProtection
        {
            CanAddOrDeleteColumns = false,
            CanAddOrDeleteRows = false,
            Password = "secret",
            EditableCellRanges = new [] { "F:S" }
        }
    }
}
```