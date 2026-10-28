# Fill Range Workbook

`UiPath.Excel.Activities.FillRange`

Fills a range in a spreadsheet with a single value or formula.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | Yes | | The range to fill (for example, `A1:B10`) |
| `Value` | Value | InArgument | `string` | | | The value to fill the range with. If it starts with an equals sign, it is written as a formula |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:FillRange
  DisplayName="Fill Range Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;A1:A10&quot;]"
  Value="[&quot;=SUM(B1:B10)&quot;]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- A literal `Value` is type-coerced like Excel would (boolean, then number, then time-of-day, then date, falling back to text) so e.g. a numeric fill is readable by `SUM`, not left as text. Numbers, times and dates are parsed using the current culture (matching what Windows/Interop's `FillRangeX` gets from Excel's own `Range.Value`) — e.g. `"1,5"` is the number 1.5 on a comma-decimal locale, not 15, and `"10:30"` is the time of day 10:30, not today's date at 10:30. Percentages (`"50%"`) and currency (`"$5"`) are not coerced to numbers, unlike Excel.
- A formula `Value` is written to the first cell and filled across `Range` the way Excel's own fill/copy-paste does — relative references (e.g. `=A1*2`) shift per cell instead of being repeated verbatim.
- A null or empty `Value` fills the range with blanks.
- A whole-column or whole-row `Range` (e.g. `A:A`) is clamped to the sheet's used range rather than filling every one of the sheet's 1,048,576 rows/16,384 columns.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:A10`) is not supported and throws `ArgumentException`.
- Throws `ArgumentException` if `Range` is null/empty.
