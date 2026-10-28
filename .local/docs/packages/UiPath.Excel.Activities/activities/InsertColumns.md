# Insert Columns Workbook

`UiPath.Excel.Activities.InsertColumns`

Inserts one or more columns into a spreadsheet before or after a reference column.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | | | The range that contains the reference column. Its first row is the header row when `HasHeaders` is on, and a column letter has to be inside it. If not specified, the used range of the sheet is used |
| `ReferenceColumn` | Reference Column | InArgument | `string` | Yes | | The column relative to which the new columns are inserted: its header name when `HasHeaders` is on, its letter otherwise (for example, `"B"`) |
| `Position` | Position | InArgument | `ColumnRelativePosition` | Yes | `Before` | Specifies whether new columns are inserted before or after the reference column |
| `NumberOfColumns` | Number Of Columns | InArgument | `int` | Yes | `1` | The number of columns to insert |
| `NewColumnName` | New Column Name | InArgument | `string` | | | The header of the new column, written in the first row of the range. When several columns are inserted, it is written in the first of them |

### Options

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `HasHeaders` | Has headers | `bool` | `False` | Specifies if the first row of the range is a header row. `ReferenceColumn` is then a header name |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Enum Reference

**`ColumnRelativePosition`**: `Before`, `After`

## XAML Example

```xml
<excel:InsertColumns
  DisplayName="Insert Columns Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  ReferenceColumn="[&quot;B&quot;]"
  Position="[ColumnRelativePosition.After]"
  NumberOfColumns="[2]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `HasHeaders` is off by default. This differs from the Windows Insert Column activity, where it is on by default: set it to `True` to keep the first row of the range.
- Existing columns at and after the insertion point shift right. Whole sheet columns are inserted, not only the cells of `Range`, which only tells where the headers are.
- `NewColumnName` is written in the first row of `Range` (of the used range when there is none).
- A header name that is not in the first row of the range, or a column letter outside the range, throws `ArgumentException` and nothing is inserted. `HasHeaders` is off by default, so existing workflows that use column letters keep working.
- Inserting a column into a table is not supported yet.
- `NumberOfColumns` must be greater than zero, otherwise `ArgumentOutOfRangeException` is thrown.
