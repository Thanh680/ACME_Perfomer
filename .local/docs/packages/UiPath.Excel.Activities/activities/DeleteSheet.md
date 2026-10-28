# Delete Sheet Workbook

`UiPath.Excel.Activities.DeleteSheet`

Deletes an existing sheet from a workbook.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | Sheet Name | InArgument | `string` | Yes | | The name of the sheet to delete |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:DeleteSheet
  DisplayName="Delete Sheet Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;OldSheet&quot;]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- Throws `ArgumentException` if `SheetName` is null/empty. Throws if the sheet does not exist, or if it is the only remaining sheet in the workbook (a workbook must always have at least one sheet).
- Throws `InvalidOperationException` if deleting the sheet would leave the workbook with no visible sheet (e.g. deleting the last visible sheet while every other sheet is hidden) — Excel itself refuses this, since a workbook saved with none fails to open.
