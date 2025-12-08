# Text alignment

Text can be aligned horizontally or vertically within a cell.
These can be set using the `AlignHorizontal` and `AlignVertical` properties on the `TableCell` class.
The valeus which can be used are defined in the `HorizontalAlignment` and `VerticalAlignment` enums in the OpenXML library.

```csharp
using Audacia.Spreadsheets;

var tableCell = new TableCell{ 
    Value = "Example text",
    AlignHorizontal = HorizontalAlignmentValues.Center,
    AlignVertical = VerticalAlignmentValues.Top
}
```
