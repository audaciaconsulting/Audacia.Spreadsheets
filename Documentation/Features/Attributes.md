# Attributes

## Creating worksheets for entities that use attributes

Usage:
```csharp
collection.ToTable();
```

or

```csharp
collection.ToWorksheet();
```

## Cell Background Colour
Allows you to set the background colour of a cell or column based on a static or dynamic value.

Fields:
- ReferenceField
- Colour

Usage:
```csharp
[CellBackgroundColour(Colour = "FFFFFF")]
public string DataProperty { get; set; }
```

or 

```csharp
public string ColourProperty { get; set; } = "FFFFFF";

[CellBackgroundColour(ReferenceField = nameof(ColourProperty))]
public string DataProperty2 { get; set; }
```


## Cell Foreground Colour
Allows you to set the foreground colour of a cell or column based on a static or dynamic value.

Fields:
- ReferenceField
- Colour

Usage:
```csharp
[CellTextColour(Colour = "FFFFFF")]
public string DataProperty { get; set; }
```

or 

```csharp
public string ColourProperty { get; set; } = "FFFFFF";

[CellTextColour(ReferenceField = nameof(ColourProperty))]
public string DataProperty { get; set; }
```

## Cell Format
For more information see the [Cell Format](./CellFormat.md) enum.

Fields:
- CellFormat

Usage:
```csharp
[CellFormat(CellFormat.Text)]
public string DataProperty { get; set; }
```

## Cell Header Name
This is a way to set the column name on the spreadsheet.
Logic will vary based on which attribute you use.

Fields:
- Name

Usage:
```csharp
[CellHeader(Name = "SomeText50")]
public string DataProperty { get; set; }
```

Alternative:
```csharp
[Display(Name = "Some Text 50")]
public string DataProperty { get; set; }
```

Output:
```csharp
Some Text 50
```

## Cell Ignore
This will mean that the column and data will not be written to the spreadsheet.

Usage:
```csharp
[CellIgnore]
public string DataProperty { get; set; }
```

Alternative:
```csharp
[IgnoreDataMember]
public string DataProperty { get; set; }
```

## Hide Header
This will set the column header text to an empty string.

Usage:
```csharp
[HideHeader]
public string DataProperty { get; set; }
```

## ID Column
This will hide the column from the spreadsheet.
The intention is that you can reuse this attribute on an entity to know it is the primary key. 

Usage:
```csharp
[IdColumn]
public string DataProperty { get; set; }
```

## Subtotal Header
This will add a subtotal cell above the column header.
Can only be used with numeric data.

Usage:
```csharp
[SubtotalHeader]
public int DataProperty { get; set; }
```
