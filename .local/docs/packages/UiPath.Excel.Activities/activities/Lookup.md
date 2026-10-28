# Lookup Workbook

`UiPath.Excel.Activities.Lookup`

Looks up a value in a vector or range and returns the corresponding value from a result vector or range (Excel `LOOKUP` function semantics).

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `LookupRange` | Lookup Range | InArgument | `string` | Yes | | The range to search in (a single row or column vector, or the result vector itself if `ResultRange` is not set) |
| `LookupValue` | Lookup Value | InArgument | `object` | Yes | | The value to search for |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `ResultRange` | Result Range | `InArgument<string>` | | The range to return a value from (a single row or column vector, with the same number of cells as `LookupRange`). If not set, `LookupRange` itself is used |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Output

| Name | Display Name | Kind | Type | Description |
|------|-------------|------|------|-------------|
| `Result` | Result | OutArgument | any (non-generic) | The value found, or `null` if the lookup value was not found. Can be bound to a variable of any type, like on the Windows activity; the value is converted when the type differs |

## XAML Example

```xml
<excel:Lookup
  DisplayName="Lookup Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  LookupRange="[&quot;A2:A100&quot;]"
  LookupValue="[&quot;Bob&quot;]"
  ResultRange="[&quot;B2:B100&quot;]"
  >
  <excel:Lookup.Result>
    <OutArgument x:TypeArguments="x:Object">[foundValue]</OutArgument>
  </excel:Lookup.Result>
</excel:Lookup>
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `Lookup` performs an **approximate match**: it returns the value corresponding to the largest entry in `LookupRange` that is less than or equal to `LookupValue`. `LookupRange` must be sorted in ascending order for correct results — this matches Excel's own `LOOKUP` function behavior.
- If `LookupValue` is smaller than every entry in `LookupRange`, or of an incompatible type (e.g. a number against a range of text), no match is found and `Result` is `null`. A number and a date **are** compatible (Excel stores dates as serial-day numbers internally), so a numeric `LookupValue` can match a date cell and vice versa.
- For an **exact-match-only** lookup, use [VLookup Workbook](VLookup.md) with `ExactMatch=True` instead.
- `LookupRange` and `ResultRange` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:A10`) is not supported and throws `ArgumentException`.
- A whole-column or whole-row `LookupRange`/`ResultRange` (e.g. `A:A`) is clamped to the sheet's used range rather than scanning every one of the sheet's 1,048,576 rows/16,384 columns.
- Throws `ArgumentException` if `LookupRange` is null/empty, if `LookupRange` or `ResultRange` is not a single row/column vector, or if `LookupRange` and `ResultRange` contain a different number of cells; throws `ArgumentNullException` if `LookupValue` is `null`.
