# Sort Range Workbook

`UiPath.Excel.Activities.SortRange`

Sorts a range by one or more columns.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | Yes | | Specifies the range to sort, in `A1:B10` format |
| `SortColumns` | Sort Columns | InArgument | `IEnumerable<SortColumnCriteria>` | Yes | | The columns to sort by (a list, an array or any sequence), in priority order. Each entry specifies a column name (or header text, if `HasHeaders` is true) and a sort direction |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `HasHeaders` | Has Headers | `bool` | `True` | Specifies if the first row of the range should be treated as a header row and excluded from sorting |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Enum Reference

**`OrderType`** (used by `SortColumnCriteria.Order`): `Ascending`, `Descending`

`SortColumnCriteria` is a plain object with two members: `ColumnName` (`string`) and `Order` (`OrderType`, default `Ascending`).

## XAML Example

```xml
<excel:SortRange
  DisplayName="Sort Range Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;A1:D100&quot;]"
  HasHeaders="True"
  SortColumns="[New List(Of SortColumnCriteria) From {New SortColumnCriteria With {.ColumnName = &quot;Amount&quot;, .Order = OrderType.Descending}}]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- When `HasHeaders` is `True`, `ColumnName` must match a header exactly — it is never interpreted as a column letter in this mode. When `False`, it must be a column letter (for example, `"B"`) within `Range`.
- The header row (if any) is excluded from the sort and never moves.
- A whole-column or whole-row `Range` (e.g. `A:C`) is clamped to the sheet's used range rather than operating over every one of the sheet's 1,048,576 rows/16,384 columns.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:C10`) is not supported and throws `ArgumentException`.
- Throws `ArgumentException` if `Range` is null/empty, if `SortColumns` is null/empty, if `SortColumns` contains a null entry or an entry whose `ColumnName` is null/empty, if `HasHeaders` is `True` and a `ColumnName` doesn't match any header, or if `HasHeaders` is `False` and a `ColumnName` is a column letter outside `Range`.
