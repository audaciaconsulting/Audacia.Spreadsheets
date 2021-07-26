# Custom Colours

This library does not support importing colours.

## Table Headers

By default table headers are a sky blue colour.
If you or your client don't like this you can replace it using a `TableHeaderStyle`.

```csharp
public class BookReport : Worksheet
{
    public BookReport(ICollection<Book> source)
    {
        var table = new Table(includeHeaders: true)
        {
            TableHeaderStyle = new TableHeaderStyle();
        };
        
        ...
    }
...
}
```

The default fill colour for a new TableHeaderStyle is White.
Colours can be defined as hex codes.

```csharp
public class TableHeaderStyle
{
    public string TextColour { get; set; } = "000000";
    public string FillColour { get; set; } = "FFFFFF";
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool HasStrike { get; set; }
    public double FontSize { get; set; } = 10d;
    public string FontName { get; set; } = "Calibri";
}
```

## Cell Colours

Cell colours can be customised during the transformation of your data to a cell.
For more information on creating worksheets see [Imports](./Imports.md).

Currently there are no extra options other than customising cell colours.
This library does not currently support making specific cells **Bold** or *Italic*.

```csharp
new TableCell("Some interesting coloured text"),
{
    FillColour = "CCCCCC",
    TextColour = "FF0000"
}
```
