# WorksheetImporter`<TRowModel`>

Worksheet importer is a generic worksheet parser, which attempts to convert each row in the worksheet to a row model object.

If an imported spreadsheet does not meet the requirements for column headers, then a single `ImportRow<TRowModel>` with validation errors will be returned.

In cases where a cell value cannot be parsed validation errors will be added to the `ImportRow<TRowModel>`.

This class can be extended to provide further validation functionality after the initial `ParseRow()` has occurred.

## Public functions

### ParseWorksheet

Iterates through and parses every row in the worksheet. Requires that the worksheet has column headers by default, this can be overriden by using `.SkipColumnHeaderMapping()` and `.MapColumn()` to define the required columns.

**Parameters:**

- Worksheet, `WorksheetBase`
- IgnoreProperties, `string[]`

**Returns**: `IEnumerable<ImportRow<TRowModel>>`

### MapColumn

Manually maps an column on the worksheet to a property on the row model.

**Parameters:**

- ColumnHeader, `string` (optional)
- PropertyExpression, `Expression<Func<TRowModel, object>>`

**Returns**: `WorksheetImporter<TRowModel>`

### SkipColumnHeaderMapping

Overrides automatic column header mapping when parsing the worksheet. This is required for parsing worksheets without column headers. When used, this function should always be called first before attempting to map columns.

**Returns**: `WorksheetImporter<TRowModel>`

## Protected properties

### ExpectedColumns

**Type:** `IDictionary<string, ProperyInfo>`

A dictionary which maps the expected column headers to properties on the imported row model type.

### SpreadsheetColumns

**Type:** `IDictionary<string, int>`

A dictionary which maps the column headers in the worksheet to field index on the TableRow.

### CurrentRow

**Type:** `TableRow`

The current TableRow being parsed by the importer.

## Protected functions

### GetColumnHeader

Gets the column header for a property.

**Parameters:**

- PropertyExpression `Expression<Func<TRowModel, object>>`

**Returns:** `string`

### GetRowNumber

Returns the ID of the current row (defaults to zero).

**Returns:** `int`

### ParseRow

Handles the parsing of the CurrentRow, should add validation errors where necessary. Can be overridden in cases where extra logic is required to parse the object.

**Parameters:**

- `out` Model, `TRowModel`

**Returns**: `IEnumerable<IImportError>`

### ValidateRow

Override this function to add further validation logic for the current row. This function will not be called if the row is already considered to be invalid after parsing.

**Parameters:**

- Model, `TRowModel`

**Returns**: `IEnumerable<IImportError>`

### TryGetBoolean

Try get cell value from the current row as a boolean.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `bool`

**Returns:** `bool`

### TryGetCell

Try get cell data from the current row. You may want to get the cell for direct access to the `cell.Value` or `cell.FillColour`.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `TableCell`

or

- ColumnHeader `string`
- `out` Value, `TableCell`

**Returns:** `bool`

### TryGetDateTime

Try get cell value from the current row as a DateTime.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `DateTime`

**Returns:** `bool`

### TryGetDateTimeOffset

Try get cell value from the current row as a DateTimeOffset.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `DateTimeOffset`

**Returns:** `bool`

### TryGetDecimal

Try get cell value from the current row as a decimal.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `decimal`

**Returns:** `bool`

### TryGetDouble

Try get cell value from the current row as a double.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `double`

**Returns:** `bool`

### TryGetEnum`<TEnum`>

Try get cell value from the current row as a typed enum.

**Parameters:**

- TEnum, `<TEnum>`
- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `TEnum`

**Returns:** `bool`

### TryGetEnum

Try get cell value from the current row as a enum object.

**Parameters:**

- EnumType, `Type`
- ValueString, `string`
- `out` Value, `object`

**Returns:** `bool`

### TryGetFloat

Try get cell value from the current row as a float.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `float`

**Returns:** `bool`

### TryGetInteger

Try get cell value from the current row as an int32.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `int`

**Returns:** `bool`

### TryGetTimeSpan

Try get cell value from the current row as a TimeSpan.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `TimeSpan`

**Returns:** `bool`

### TryGetString

Try get cell value from the current row as a string.

**Parameters:**

- PropertyExpression, `Expression<Func<TRowModel, object>>`
- `out` Value, `string`

or

- ColumnHeader `string`
- `out` Value, `string`

**Returns:** `bool`

### TryParseDateTime

Attempts to parse a DateTime from a string using a predefined set of datetime formats.

**Parameters:**

- ValueString, `string`
- `out` Value, `DateTime`

**Returns:** `bool`

### TryParseDateTimeOffset

Attempts to parse a DateTimeOffset from a string using a predefined set of datetime formats.

**Parameters:**

- ValueString, `string`
- `out` Value, `DateTimeOffset`

**Returns:** `bool`

### TryParseEnum

Attempts to parse a Enum from a string.

**Parameters:**

- EnumType, `Type`
- ValueString, `string`
- `out` Value, `object`

**Returns:** `bool`

### TryParseEnum`<TEnum`>

Attempts to parse a Enum from a string.

**Parameters:**

- TEnum, `<TEnum>`
- ValueString, `string`
- `out` Value, `TEnum`

**Returns:** `bool`

### TryParseTimeSpan

Attempts to parse a TimeSpan from a string.

**Parameters:**

- ValueString, `string`
- `out` Value, `TimeSpan`

**Returns:** `bool`
