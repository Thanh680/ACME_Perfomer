# Find First or Last Data Row Workbook

`UiPath.Excel.Activities.FindFirstLastDataRow`

Returns the index of the first row containing data, and the index of the last populated row (or the first empty row after it), in a column of a range. The cross-platform counterpart of Find First or Last Data Row (`FindFirstLastDataRowX`).

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | | | Specifies the range to search. If not specified, the whole used range of the sheet is used |
| `ColumnName` | Column name | InArgument | `string` | Yes | | Column to search for data values: its header name when `HasHeaders` is `True`, its letter otherwise |
| `FirstRowOffset` | First row offset | InArgument | `int` | | `0` | Adds the number specified to the row index of the first row containing data. It counts rows (visible rows when `VisibleRowsOnly` is on), not populated cells, so the result can be a blank row. Enables scenarios such as "the row 2 rows below the first data row" |
| `LastRowOffset` | Last row offset | InArgument | `int` | | `0` | Subtracts the number specified from the row index of the last row containing data. It counts rows (visible rows when `VisibleRowsOnly` is on), not populated cells, so the result can be a blank row. Enables scenarios such as "the row 3 rows above the last data row" |
| `BlankRowsToSkip` | Blank rows to skip | InArgument | `int` | | `1` (when not set) | Consecutive blank rows allowed in the data prior to determining that the data in the range has ended |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `HasHeaders` | Has headers | `bool` | `True` | Indicates if the first row of the range is a header row. The data starts below it |
| `ConfigureLastRowAs` | Configure last row as | `InArgument<LastRowConfiguration>` | `LastPopulatedRow` | Whether the last row is the last populated row, or the first empty row after it. Does not affect the first row |
| `VisibleRowsOnly` | Visible rows only | `bool` | `False` | Uses only the visible rows, taking into account the filters applied and ignoring the hidden rows |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Output

| Name | Display Name | Kind | Type | Description |
|------|-------------|------|------|-------------|
| `FirstRowIndex` | First row index | OutArgument | `int` | The 1-based worksheet row number of the first row containing data. `-1` if no data is found. This is the absolute row number in the sheet, not an offset relative to `Range` |
| `LastRowIndex` | Last row index | OutArgument | `int` | The 1-based worksheet row number of the last populated row (or the first empty row, depending on `ConfigureLastRowAs`). `-1` if no data is found, or `0` when `ConfigureLastRowAs` is `FirstEmptyRow`. This is the absolute row number in the sheet, not an offset relative to `Range` |

### Enum Reference

**`LastRowConfiguration`**: `LastPopulatedRow`, `FirstEmptyRow`

## XAML Example

```xml
<excel:FindFirstLastDataRow
  DisplayName="Find First or Last Data Row Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;A1:D1000&quot;]"
  ColumnName="[&quot;Name&quot;]"
  HasHeaders="True"
  BlankRowsToSkip="[5]"
  ConfigureLastRowAs="[LastRowConfiguration.FirstEmptyRow]"
  FirstRowIndex="[firstRow]"
  LastRowIndex="[lastRow]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- Read-only — does not modify the workbook.
- Only the cells of `ColumnName` are searched. A cell holds data when it is not empty and, for text, not blank or whitespace; numbers, dates, booleans and errors are data, and a formula counts through the value it calculates.
- The search stops at the first run of more than `BlankRowsToSkip` consecutive blank rows. With the default of `1`, a gap of two blank rows ends the data: the first row must come within that many blank rows from the start of the data, and the last row is the last populated row before such a gap.
- The offsets are applied last. The first row moves down by `FirstRowOffset`; the last row moves up by `LastRowOffset` (for `FirstEmptyRow` the result is the row after the last populated one, moved up by the offset). A result outside the sheet, or no data at all, gives `-1` (and `0` for the last row when `ConfigureLastRowAs` is `FirstEmptyRow`).
- With `VisibleRowsOnly`, hidden rows (for example the rows hidden by `Filter Range Workbook`) are ignored: they do not count as blank rows and the offsets count visible rows only.
- Also works for a whole-sheet search (no `Range`): the rows searched are those of the sheet's used range and `ColumnName` can be any column letter. A whole-column or whole-row `Range` (e.g. `A:A`) is clamped to the sheet's used range, starting from row 1 for a column.
- Tables, named ranges and pivot tables are not available as a range reference; pass the table or named range address as `Range` instead (a table includes its header row, so keep `HasHeaders` as `True`).
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:A10`) is not supported and throws `ArgumentException`. An unknown header, or a column letter outside `Range`, throws `ArgumentException`.
