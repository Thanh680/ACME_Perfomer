# Delete Columns Workbook

`UiPath.Excel.Activities.DeleteColumns`

Deletes one or more consecutive columns from a spreadsheet starting at a specified column.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | | | The range that contains the column. Its first row is the header row when `HasHeaders` is on, and a column letter has to be inside it. If not specified, the used range of the sheet is used |
| `StartColumn` | Start Column | InArgument | `string` | Yes | | The column to start deleting from: its header name when `HasHeaders` is on, its letter otherwise (for example, `"B"`) |
| `NumberOfColumns` | Number Of Columns | InArgument | `int` | Yes | `1` | The number of columns to delete |

### Options

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `HasHeaders` | Has headers | `bool` | `False` | Specifies if the first row of the range is a header row. `StartColumn` is then a header name |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:DeleteColumns
  DisplayName="Delete Columns Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  StartColumn="[&quot;C&quot;]"
  NumberOfColumns="[1]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `HasHeaders` is off by default. This differs from the Windows Delete Column activity, where it is on by default: set it to `True` to keep the first row of the range.
- The whole sheet column is deleted, not only the cells of `Range`; the range only tells where the headers are (or which letters are valid). Deleting columns of a table is not supported yet.
- A header name that is not in the first row of the range, or a column letter outside the range, throws `ArgumentException`.
- Columns after the deleted range shift left to fill the gap.
- `NumberOfColumns` must be greater than zero, otherwise `ArgumentOutOfRangeException` is thrown.
