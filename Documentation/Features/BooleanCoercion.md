# Parsing booleans from strings

```csharp
using Audacia.Spreadsheets;

var myBool = Bool.Parse("Yes");
var success = Bool.TryParse("False", out var myOtherBool);
```

## Supported Values

All value checks are case insensitive.
The method will trim the string for you when checking.

### True Values

- True
- Yes
- Y
- 1

### False Values

- False
- No
- N
- 0
