# Protect Sheet Workbook

`UiPath.Excel.Activities.ProtectSheet`

Protects the current sheet, optionally allowing a set of actions even while protected.

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
| `AllowedPermissions` | Allowed Permissions | `InArgument<ProtectSheetAdditionalPermissions>` | `None` (when not set) | Actions that remain allowed even while the sheet is protected |
| `SheetPassword` | Protection Password | `InArgument<string>` | | The password that protects the sheet and is required to unprotect it. If not set, the sheet is protected without a password. Distinct from `Password` below, which opens the workbook *file* |
| `Password` | Password | `InArgument<string>` | | The password of the workbook *file*, if necessary — not related to sheet protection |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Enum Reference

**`ProtectSheetAdditionalPermissions`** (flags): `None`, `AllowDeletingColumns`, `AllowDeletingRows`, `DrawingObjects`, `Scenarios`, `AllowFiltering`, `AllowFormattingCells`, `AllowFormattingColumns`, `AllowFormattingRows`, `AllowInsertingColumns`, `AllowInsertingHyperlinks`, `AllowInsertingRows`, `AllowSorting`, `AllowUsingPivotTables`

## XAML Example

```xml
<excel:ProtectSheet
  DisplayName="Protect Sheet Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  SheetPassword="[sheetPassword]"
  AllowedPermissions="[ProtectSheetAdditionalPermissions.AllowSorting Or ProtectSheetAdditionalPermissions.AllowFiltering]" />
```

`sheetPassword` above is a workflow variable (e.g. sourced from an Orchestrator asset or credential activity) — avoid hard-coding a password as a literal string in a source-controlled workflow file.

## Errors

| Condition | Exception |
|-----------|-----------|
| The sheet is already protected | `ExcelException` — "The {SheetName} sheet is already protected. Unprotect it before setting new permissions." |
| Workbook file is `.xls` | `NotSupportedException` |

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- `ProtectSheetAdditionalPermissions` is a flags enum — combine values with `Or` (e.g. `AllowSorting Or AllowFiltering`) to allow multiple actions.
- See also [Unprotect Sheet Workbook](UnprotectSheet.md).
