# Insert Sheet Workbook

`UiPath.Excel.Activities.InsertSheet`

Inserts a new sheet into a workbook, optionally at a specific position.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `NewSheetName` | New Sheet Name | InArgument | `string` | Yes | | The name of the sheet to create |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `Position` | Position | `InArgument<int?>` | | The 1-based position to insert the new sheet at. Must be at least 1 (throws `ArgumentOutOfRangeException` otherwise). A value greater than the current sheet count appends at the end. If not specified, the sheet is appended at the end |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

## XAML Example

```xml
<excel:InsertSheet
  DisplayName="Insert Sheet Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  NewSheetName="[&quot;Summary&quot;]"
  Position="[1]" />
```

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- Unlike most other Portable Excel activities, this activity has no `SheetName` property — it creates the sheet itself and does not operate against a "current" sheet.
- Throws `ArgumentException` if a sheet with the given name already exists, or if `NewSheetName` is null/empty.
- Throws `ArgumentOutOfRangeException` if `Position` is less than 1.
