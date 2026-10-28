# Unprotect Sheet Workbook

`UiPath.Excel.Activities.UnprotectSheet`

Removes protection from the current sheet.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `SheetPassword` | Protection Password | `InArgument<string>` | | The password that was used to protect the sheet, if any. Distinct from `Password` below, which opens the workbook *file* |
| `Password` | Password | `InArgument<string>` | | The password of the workbook *file*, if necessary — not related to sheet protection |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:UnprotectSheet
  DisplayName="Unprotect Sheet Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  SheetPassword="[sheetPassword]" />
```

`sheetPassword` above is a workflow variable (e.g. sourced from an Orchestrator asset or credential activity) — avoid hard-coding a password as a literal string in a source-controlled workflow file.

## Errors

| Condition | Exception |
|-----------|-----------|
| `SheetPassword` doesn't match the sheet's protection password | `ExcelException` — "The provided password is incorrect." |
| The sheet is password-protected and `SheetPassword` is not set | `ExcelException` — "Password is required." |
| Workbook file is `.xls` | `NotSupportedException` |

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- Unprotecting a sheet that is not currently protected is a no-op, not an error.
- See also [Protect Sheet Workbook](ProtectSheet.md).
