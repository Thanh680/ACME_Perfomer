# Auto Fit Workbook

`UiPath.Excel.Activities.AutoFit`

Autofits the columns and/or rows of a range to their contents.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | | | The range to autofit. If not specified, the whole used range of the sheet is used |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `FitColumns` | Fit Columns | `bool` | `True` | Autofit column widths to their contents |
| `FitRows` | Fit Rows | `bool` | `True` | Autofit row heights to their contents |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:AutoFit
  DisplayName="Auto Fit Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;A1:D50&quot;]"
  FitColumns="True"
  FitRows="True" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- Column width and row height are whole-column/whole-row properties in Excel — autofitting based on a narrower `Range` only affects the sizing *calculation* (which cells are considered), not the columns/rows actually resized (the full column/row is still resized).
- A whole-column or whole-row `Range` (e.g. `A:A`) is clamped to the sheet's used range for the sizing calculation, rather than considering every one of the sheet's 1,048,576 rows/16,384 columns.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:A10`) is not supported and throws `ArgumentException`.
