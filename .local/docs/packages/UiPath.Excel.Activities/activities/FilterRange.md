# Filter Range Workbook

`UiPath.Excel.Activities.FilterRange`

Filters a range by the values of one column (Basic: a list of allowed values; Advanced: one or two conditions), or clears the filters applied on it. The cross-platform counterpart of Filter (`FilterX`).

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | | | The range to filter. The first row of the range is the header row. If not specified, the whole used range of the sheet is used |
| `ColumnName` | Column Name | InArgument | `string` | Yes (when not clearing) | | The header name of the column to filter on. When `ClearFilter` is `True`, only the filter of this column is cleared; if empty, every filter of the sheet is cleared |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `ClearFilter` | Clear any existing filter | `bool` | `False` | Clears the existing filters instead of applying one |
| `UseAdvancedFilter` | Use advanced filter | `bool` | `False` | Switches between the Basic and the Advanced filter. Designer only; it is persisted as `FilterArgument.Mode` |
| `BasicFilter` | Basic filter | list of values | | Rows whose value in `ColumnName` equals any of the listed values are kept. Designer only; it is persisted in `FilterArgument` |
| `AdvancedFilter` | Advanced filter | one or two conditions | | Conditions made of an operator and a value, combined with `And` / `Or`. Designer only; it is persisted in `FilterArgument` |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

The `FilterArgument` property holds the Basic or Advanced filter and is not shown in the designer, exactly like in `FilterX`.

### Operators (Advanced filter)

`<`, `>`, `<=`, `>=`, `=`, `!=`, Is empty, Is not empty, Starts with, Ends with, Contains, Does not start with, Does not end with, Does not contain.

- Numbers and dates are compared as such, text is compared case-insensitively. A text operator only matches text cells; its negations also match numbers and empty cells.
- `*` (any characters), `?` (one character) and `~` (escape) are wildcards in the value of the `=`, `!=` and text operators.
- Is empty and Is not empty take no value; every other operator requires one.
- When two conditions are given, `And` / `Or` combines them. The operator of an unused condition is `None`.

## XAML Example

```xml
<excel:FilterRange
  DisplayName="Filter Range Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;A1:D100&quot;]"
  ColumnName="[&quot;Amount&quot;]">
  <excel:FilterRange.FilterArgument>
    <excelfilter:FilterArgument Mode="Advanced">
      <excelfilter:FilterArgument.AdvancedFiltersArgument>
        <excelfilter:AdvancedFilterArgument LogicalOperator="And">
          <excelfilter:AdvancedFilterArgument.Condition1>
            <excelfilter:AdvancedFilterCondition Operator="GT" Value="[&quot;50&quot;]" />
          </excelfilter:AdvancedFilterArgument.Condition1>
          <excelfilter:AdvancedFilterArgument.Condition2>
            <excelfilter:AdvancedFilterCondition Operator="LT" Value="[&quot;200&quot;]" />
          </excelfilter:AdvancedFilterArgument.Condition2>
        </excelfilter:AdvancedFilterArgument>
      </excelfilter:FilterArgument.AdvancedFiltersArgument>
    </excelfilter:FilterArgument>
  </excel:FilterRange.FilterArgument>
</excel:FilterRange>
```

`excelfilter` maps to `clr-namespace:UiPath.Excel.Activities.Business.Filter;assembly=UiPath.Excel.Activities`. To clear the filter of a column, set `ClearFilter="True"` and `ColumnName`; to clear every filter of the sheet, set `ClearFilter="True"` with no `ColumnName`.

## Notes

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- An Advanced filter with `GT`, `GTE`, `LT` or `LTE` on a number or a date compares the values of the cells, not the text a number format shows (a column showing `1` for 1.2 and 1.4 is still split by "greater than 1.3"). The other operators, `EQ` and `NOTEQ` included, are saved as the list of the displayed values that match, as Excel does for a value filter, so two cells that show the same text are kept or hidden together.
- Filters on several columns of the same range accumulate. A whole-column address such as `A:C` is resolved against the rows used now, so the filters stay together when rows are appended.
- The first row of `Range` is always the header row and is never hidden. `ColumnName` is matched against the header names (case-insensitive); it is never read as a column letter.
- The filter is stored in the workbook as an AutoFilter on `Range`, so Excel shows the filter buttons and the hidden rows. A sheet holds a single AutoFilter:
  - Filtering another column of the same `Range` keeps the existing filters, so the criteria of different columns combine with AND.
  - Filtering a column again replaces its previous filter.
  - Filtering on a different `Range` removes the existing filter, showing its rows again.
- The filters are stored as the list of displayed values that match, so rows added later are not matched until the filter is applied again. Rows hidden by hand are shown again when they match.
- Basic values are compared case-insensitively against the text displayed in the cell (a date as `1/5/2020`, a number as `1,234.50`, an empty cell as an empty string).
- Clearing a column that has no filter, or clearing when the sheet has no filter, does nothing. Clearing a column fails with `InvalidOperationException` when the existing filter is on another `Range`.
- A whole-column or whole-row `Range` (e.g. `A:B`) is clamped to the sheet's used range rather than operating over every one of the sheet's 1,048,576 rows/16,384 columns.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A:B`) is not supported and throws `ArgumentException`. `Range` must not overlap an Excel table (`InvalidOperationException`).
- Validation: `ColumnName` and the filter are required unless `ClearFilter` is `True`. An unmatched header throws `ArgumentException`.
