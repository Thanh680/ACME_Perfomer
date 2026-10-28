# VLookup Workbook

`UiPath.Excel.Activities.VLookup`

Looks up a value in the first column of a table range and returns a value from the given column index (Excel `VLOOKUP` function).

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `TableRange` | Table Range | InArgument | `string` | Yes | | The table range to search. The search happens in its first column |
| `LookupValue` | Lookup Value | InArgument | `object` | Yes | | The value to search for |
| `ColumnIndex` | Column Index | InArgument | `int` | Yes | | The 1-based column index within `TableRange` to return a value from |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `ExactMatch` | Exact Match | `bool` | `True` | `True` for an exact match; `False` for an approximate match against sorted data (nearest value less than or equal to the lookup value) |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Output

| Name | Display Name | Kind | Type | Description |
|------|-------------|------|------|-------------|
| `Result` | Result | OutArgument | any (non-generic) | The value found, or `null` if the lookup value was not found. Can be bound to a variable of any type, like on the Windows activity; the value is converted when the type differs |

## XAML Example

```xml
<excel:VLookup
  DisplayName="VLookup Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  TableRange="[&quot;A1:C100&quot;]"
  LookupValue="[&quot;Bob&quot;]"
  ColumnIndex="[3]"
  ExactMatch="True"
  >
  <excel:VLookup.Result>
    <OutArgument x:TypeArguments="x:Object">[foundValue]</OutArgument>
  </excel:VLookup.Result>
</excel:VLookup>
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- With `ExactMatch=False`, `TableRange`'s first column must be sorted in ascending order — matches Excel's own `VLOOKUP` behavior for approximate matches.
- `ColumnIndex` is 1-based and relative to `TableRange` (column 1 is the range's first column, the one being searched).
- `TableRange` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:C100`) is not supported and throws `ArgumentException`.
- A whole-column or whole-row `TableRange` (e.g. `A:C`) is clamped to the sheet's used range rather than scanning every one of the sheet's 1,048,576 rows/16,384 columns.
- Throws `ArgumentException` if `TableRange` is null/empty, `ArgumentOutOfRangeException` if `ColumnIndex` is less than 1 or greater than the number of columns in `TableRange`, and `ArgumentNullException` if `LookupValue` is `null`.
