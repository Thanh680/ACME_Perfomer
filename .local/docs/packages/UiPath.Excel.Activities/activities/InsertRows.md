# Insert Rows Workbook

`UiPath.Excel.Activities.InsertRows`

Inserts one or more rows into a spreadsheet at a specified position.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | | | The range to insert rows into. If not specified, the used range of the sheet is used |
| `NumberOfRows` | Number Of Rows | InArgument | `int` | Yes | `1` | The number of rows to insert |
| `Position` | Position | InArgument | `InsertRowPosition` | Yes | `End` | Specifies where the new rows should be inserted (Start, End, or a specific row index) |
| `SpecificRowIndex` | Specific Row Index | InArgument | `int?` |  |  | The 1-based row index, relative to the start of `Range`, at which to insert rows. Required when `Position` is `Specific`; only shown in the designer then |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `HasHeaders` | Has Headers | `bool` | `False` | Specifies if the first row of the range is a header row. The header row stays on top: rows inserted at the start go below it |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Enum Reference

**`InsertRowPosition`**: `Start`, `End`, `Specific`

## Valid Configurations

- **Position = Start / End**: `SpecificRowIndex` is not used.
- **Position = Specific**: `SpecificRowIndex` must be set to a 1-based row index; the activity throws `ArgumentException` if it is left unset, and `ArgumentOutOfRangeException` if it is below 1 or beyond the last row of the range (nothing is inserted).

## XAML Example

```xml
<excel:InsertRows
  DisplayName="Insert Rows Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Position="[InsertRowPosition.Specific]"
  SpecificRowIndex="[3]"
  NumberOfRows="[2]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `HasHeaders` is off by default. This differs from the Windows Insert Rows activity, where it is on by default: set it to `True` to keep the first row of the range.
- Existing rows at and below the insertion point are shifted down.
- `NumberOfRows` must be greater than zero, otherwise `ArgumentOutOfRangeException` is thrown.
- Without `Range`, whole sheet rows are inserted relative to the used range of the sheet. With `Range`, only the cells of the range's columns shift down, like the Windows `InsertRowsX` activity; the rows below the range in those columns move down, and other columns are untouched.
- `Position = End` inserts immediately after the last row of the range. `Start` and `Specific` positions are counted from the first row of the range, below the header row when `HasHeaders` is on.
- `HasHeaders` is off by default: `Start` inserts at the very first row of the range. Turn it on to insert below the header row.
- Inserting into a table is not supported yet.
