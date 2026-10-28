# Duplicate Sheet Workbook

`UiPath.Excel.Activities.DuplicateSheet`

Duplicates an existing sheet in a workbook under a new name.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | Sheet Name | InArgument | `string` | Yes | | The name of the existing sheet to duplicate |
| `NewSheetName` | New Sheet Name | InArgument | `string` | Yes | | The name for the duplicated sheet |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:DuplicateSheet
  DisplayName="Duplicate Sheet Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Template&quot;]"
  NewSheetName="[&quot;Template (Copy)&quot;]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- The new sheet is placed immediately after the source sheet, and contains a full copy of its data.
- Throws `ArgumentException` if `SheetName` or `NewSheetName` is null/empty. Throws if the source sheet does not exist or the target name already exists.
