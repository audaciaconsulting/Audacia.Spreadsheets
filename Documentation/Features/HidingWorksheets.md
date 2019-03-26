# Hiding Worksheets

You can set worksheets as hidden in excel.
This can be done by calling the Hide() method or setting the Visibility property on a worksheet.

```csharp
using Audacia.Spreadsheets;

class MyWorksheet : Worksheet
{
    public MyWorksheet() 
    {
        Hide();
    }
}
```

Sometimes you may want to prevent the user from seeing hidden worksheets.
Setting this sets the visibility of the worksheet to _Very Hidden_. 

```csharp
using Audacia.Spreadsheets;

class MyWorksheet : Worksheet
{
    public MyWorksheet() 
    {
        Hide(completelyHidden: true);
    }
}
```
