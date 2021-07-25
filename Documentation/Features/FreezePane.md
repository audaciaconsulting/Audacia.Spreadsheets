# Freeze Panes

You can add freeze panes to worksheets.

**Be aware**, on the current version of the library if you have **more than one table** and a freeze pane, this will likely corrupt your worksheet.

## Creating a freeze pane

If you create a new Freeze Pane it will automatically be setup to freeze the top row.

```csharp
public class HouseReport : Worksheet
{
    public HouseReport(...)
    {
        ...
        
        FreezePane = new FreezePane()
    }
}
```

Change the starting cell to move the freeze pane about.
Change the FrozenColumn & FrozenRow properties to change the size of the frozen pane.

```csharp
public class BookReport : Worksheet
{
    public BookReport(...)
    {
        ...
        
        FreezePane = new FreezePane
        {
            StartingCell = "A3",
            FrozenRows = 2D
        }
    }
}
```
