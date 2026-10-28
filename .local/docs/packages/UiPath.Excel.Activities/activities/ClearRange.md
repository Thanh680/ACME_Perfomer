# Clear Range Workbook

`UiPath.Excel.Activities.ClearRange`

Clears the contents, and optionally the formatting, of a range without shifting cells.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | Yes | | The range whose contents (and optionally formatting) will be cleared |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `HasHeaders` | Has headers | `bool` | `False` | Specifies if the first row of the range is a header row, which is not cleared. A range of a single row is left as it is |
| `ClearFormatting` | Clear Formatting | `bool` | `True` | Also clears the formatting of the range, not just its contents. On by default, as on the Windows activity |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:ClearRange
  DisplayName="Clear Range Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;A1:D10&quot;]"
  HasHeaders="True"
  ClearFormatting="True" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `ClearFormatting` is on by default and `HasHeaders` off, like the Windows Clear Range activity. With `HasHeaders`, the header row is the first row of `Range` (of the used range for a whole-column address).
- Unlike a delete/shift operation, cells outside `Range` are never moved — this only clears what's inside the range.
- A whole-column or whole-row `Range` (e.g. `A:A`) is clamped to the sheet's used range rather than operating over every one of the sheet's 1,048,576 rows/16,384 columns.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:A10`) is not supported and throws `ArgumentException`.
- Throws `ArgumentException` if `Range` is null/empty.
