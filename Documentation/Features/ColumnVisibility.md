# Column visibility

Columns can be set to hidden by setting the `IsHidden` flag on the `TableColumn` to true.
This will function exactly as hiding a column in Excel, so the column can stil lbe made visible by the user.
`IsHidden` is false by default.

```csharp
using Audacia.Spreadsheets;

var tableColumn = new TableColumn("Example Column") { IsHidden = true };
```
