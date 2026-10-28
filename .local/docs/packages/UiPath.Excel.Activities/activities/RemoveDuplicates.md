# Remove Duplicates Workbook

`UiPath.Excel.Activities.RemoveDuplicates`

Removes duplicate rows from a range, comparing all columns or a specific subset.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | | | The range to remove duplicate rows from. If not specified, the whole used range of the sheet is used |
| `Mode` | Compare | InArgument | `ColumnsCompare` | Yes | `AllColumns` | Specifies whether all columns or only a specific subset of columns should be compared when looking for duplicate rows |
| `ColumnsToCompare` | Columns To Compare | InArgument | `IEnumerable<string>` |  |  | The column names, as a list, an array or any sequence (or letters, if `HasHeaders` is false) to compare when `Mode` is set to `IndividualColumns` |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `HasHeaders` | Has Headers | `bool` | `True` | Specifies if the first row of the range should be treated as a header row and excluded from the duplicate search |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Output

| Name | Display Name | Kind | Type | Description |
|------|-------------|------|------|-------------|
| `RemovedRowsCount` | Removed Rows Count | OutArgument | `int` | The number of duplicate rows that were removed |

### Enum Reference

**`ColumnsCompare`**: `IndividualColumns`, `AllColumns`

## Valid Configurations

- **Mode = AllColumns**: every column in `Range` is compared; `ColumnsToCompare` is ignored.
- **Mode = IndividualColumns**: only the columns listed in `ColumnsToCompare` are compared. `ColumnsToCompare` must specify at least one column; the activity throws `ArgumentException` if it is left null/empty.
- When `HasHeaders` is `True`, each entry in `ColumnsToCompare` must match a header exactly — it is never interpreted as a column letter in this mode; a non-matching entry throws `ArgumentException`. When `False`, each entry must be a column letter within `Range`, or the activity throws `ArgumentException`.

## XAML Example

```xml
<excel:RemoveDuplicates
  DisplayName="Remove Duplicates Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;A1:D100&quot;]"
  Mode="[ColumnsCompare.IndividualColumns]"
  ColumnsToCompare="[New String() {&quot;Email&quot;}]"
  HasHeaders="True"
  RemovedRowsCount="[removedCount]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- The first occurrence of each unique row is kept; later duplicates are removed and rows below shift up. When `Range` covers the whole used range, row visibility (hidden/visible) travels with the shift. When `Range` is narrower than the sheet, row visibility is left untouched — matching Excel's own behavior, since a hidden/visible row is a whole-row property that can't be scoped to a subset of columns.
- A whole-column or whole-row `Range` (e.g. `A:A`) is clamped to the sheet's used range rather than operating over every one of the sheet's 1,048,576 rows/16,384 columns.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:A10`) is not supported and throws `ArgumentException`.
