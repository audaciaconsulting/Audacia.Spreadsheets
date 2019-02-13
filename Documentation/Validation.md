# Validation
Some built in validation models, you don't have to use them.
It's up to you to write your own validation logic.

All validation models implement `IImportError`.
```csharp
public interface IImportError
{
    string GetMessage();
}
```

## Sheet Validation
- **MissingWorksheetError** - For one or more missing worksheets.

## Column Validation
- **DuplicateColumnError** - For when there are multiple columns with the same name.

- **MissingColumnError** - For one or more missing columns.

## Row Validation

### Groups

- **DuplicateKeyError** - For when multiple of the same unique constraint are imported.

- **RecordExistsError** - For when a unique record is already in the system.

- **RecordMissingError** - For when the referenced record doesn't exist in the system.

- **RecordAssociationError** - For when the referenced record cannot be associated.

### Fields

- **FieldParseError** - For when a cell value cannot be parsed.

- **FieldValidationError** - For row validation errors.