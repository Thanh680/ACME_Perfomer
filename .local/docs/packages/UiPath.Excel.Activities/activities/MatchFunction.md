# Match Function Workbook

`UiPath.Excel.Activities.MatchFunction`

Returns the relative position of a value within a range (Excel `MATCH` function).

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | Yes | | The range to search within |
| `LookupValue` | Lookup Value | InArgument | `object` | Yes | | The value to search for |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `MatchType` | Match Type | `InArgument<MatchType>` | `LargestValueLessOrEqual` | Specifies how the lookup value is matched against the range |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Output

| Name | Display Name | Kind | Type | Description |
|------|-------------|------|------|-------------|
| `Position` | Position | OutArgument | `int` | The 1-based relative position of the match within `Range`, or `-1` if not found, like the Windows activity |

### Enum Reference

**`MatchType`**: `LargestValueLessOrEqual` (= 1, requires `Range` sorted ascending), `ExactlyEqual` (= 0, `Range` can be in any order), `SmallestValueGreaterOrEqual` (= -1, requires `Range` sorted descending)

## XAML Example

```xml
<excel:MatchFunction
  DisplayName="Match Function Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;A2:A100&quot;]"
  LookupValue="[&quot;Bob&quot;]"
  MatchType="[MatchType.ExactlyEqual]"
  Position="[matchPosition]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `Position` is 1-based and relative to `Range` (not an absolute sheet row/column number) — e.g. a match on the 3rd cell of `Range` returns `3`, regardless of where `Range` starts on the sheet.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A2:A100`) is not supported and throws `ArgumentException`.
- A whole-column or whole-row `Range` (e.g. `A:A`) is clamped to the sheet's used range rather than scanning every one of the sheet's 1,048,576 rows/16,384 columns.
- Throws `ArgumentException` if `Range` is null/empty, and `ArgumentNullException` if `LookupValue` is `null`.
