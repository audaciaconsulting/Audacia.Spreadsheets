# Locking a worksheet

You can add WorksheetProtection to worksheets.

By default, Excel marks all cells as 'Locked' by default. You can see this behaviour by opening an Excel file > right-click a cell > Format Cell > Protection.

The protection does not effect the document unless there is a password protecting the sheet.

To protect the entire sheet, add this code to your worksheet:

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
        }
    }
}
```

If you need particular cells to be editable, this is done at a `TableCell` level.

When creating a `TableCell` you can use the `IsEditable` property to decide if a particular cell should be editable.