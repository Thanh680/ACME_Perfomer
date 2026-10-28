# Delete Rows Workbook

`UiPath.Excel.Activities.DeleteRows`

Deletes rows from a spreadsheet: a specific range, all visible rows, all hidden rows, or duplicate rows.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Option` | Option | InArgument | `DeleteRowsOption` | Yes | `Specific` | Specifies which rows to delete: a specific range, all visible rows, all hidden rows, or duplicate rows |
| `RowPositions` | At position | InArgument | `string` |  |  | The rows to delete, relative to the start of `Range`: a comma separated list of row numbers and ranges, for example `3, 5-7, 10`. Required when `Option` is `Specific`, and only shown in the designer then |
| `Range` | Range | InArgument | `string` | | | The range to delete rows from. If not specified, the whole used range of the sheet is used |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `HasHeaders` | Has Headers | `bool` | `False` | Specifies if the first row of the range is a header row and should never be deleted |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Output

| Name | Display Name | Kind | Type | Description |
|------|-------------|------|------|-------------|
| `DeletedRowsCount` | Rows Deleted | OutArgument | `int` | The number of rows that were actually deleted |

### Enum Reference

**`DeleteRowsOption`**: `Specific`, `Visible`, `Hidden`, `Duplicates`

## Valid Configurations

- **Option = Specific**: `RowPositions` must be set. The positions are 1-based and relative to the start of `Range` (or the whole used range). A missing or malformed list throws `ArgumentException` (and is a validation error in the designer when it is a literal); a position below 1 or beyond the last row of the range throws `ArgumentOutOfRangeException`, and nothing is deleted. Blocks may be given in any order, may overlap and may be reversed (`10-8` is `8-10`); every row is deleted and counted once. When `HasHeaders` is `True`, the header row is skipped even if it is listed.
- **Option = Visible**: deletes every non-hidden row within `Range` (or the whole used range).
- **Option = Hidden**: deletes every hidden row within `Range` (or the whole used range).
- **Option = Duplicates**: deletes rows that duplicate an earlier row (comparing all columns) within `Range` (or the whole used range).
- When `HasHeaders` is `True`, the first row of the range is never deleted, regardless of `Option`.

## XAML Example

```xml
<excel:DeleteRows
  DisplayName="Delete Rows Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Option="[DeleteRowsOption.Specific]"
  Range="[&quot;A1:D50&quot;]"
  RowPositions="[&quot;2, 5-7, 10&quot;]"
  HasHeaders="True"
  DeletedRowsCount="[deletedCount]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `HasHeaders` is off by default. This differs from the Windows Delete Rows activity, where it is on by default: set it to `True` to keep the first row of the range.
- Rows below a deleted row shift up to fill the gap. When `Range` covers the whole used range, row visibility (hidden/visible) travels with the shift. When `Range` is narrower than the sheet, row visibility is otherwise left untouched — a hidden/visible row is a whole-row property that can't be scoped to a subset of columns. The one exception is **Option = Hidden**: each deleted row is unhidden right before it's deleted, so the data that shifts up into its position isn't left hidden; this also unhides any other data on that same row outside `Range`.
- A whole-column or whole-row `Range` (e.g. `A:A`) is clamped to the sheet's used range rather than operating over every one of the sheet's 1,048,576 rows/16,384 columns.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:A10`) is not supported and throws `ArgumentException`.
