# Rename Sheet Workbook

`UiPath.Excel.Activities.RenameSheet`

Renames an existing sheet in a workbook.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | Sheet Name | InArgument | `string` | Yes | | The name of the existing sheet to rename |
| `NewSheetName` | New Sheet Name | InArgument | `string` | Yes | | The new name for the sheet |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:RenameSheet
  DisplayName="Rename Sheet Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  NewSheetName="[&quot;Summary&quot;]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `SheetName` here identifies the sheet to act on — this activity does not use the "select or create a current sheet" `SheetName` semantics found on most other Portable Excel activities.
- Throws `ArgumentException` if `SheetName` or `NewSheetName` is null/empty. Throws if the target name already exists or the source sheet does not exist.
