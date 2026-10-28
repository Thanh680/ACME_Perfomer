# Format Range Workbook

`UiPath.Excel.Activities.FormatRange`

Applies number, date, time, alignment, and font settings to a range.

**Package:** `UiPath.Excel.Activities`

## Properties

### Input

| Name | Display Name | Kind | Type | Required | Default | Description |
|------|-------------|------|------|----------|---------|-------------|
| `WorkbookPath` | File (local path) | InArgument | `string` | | | The full local path of the workbook. Alternative to `WorkbookPathResource` — use one or the other |
| `WorkbookPathResource` | File | InArgument | `IResource` | | | A resource reference (e.g. a local file, remote file, or abstract resource). Alternative to `WorkbookPath` — use one or the other |
| `SheetName` | SheetName | InArgument | `string` | Yes | | The name of the sheet from the workbook |
| `Range` | Range | InArgument | `string` | Yes | | The range to format |

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `Category` | Category | `InArgument<CellFormat>` | | The number format category to apply. Leave unset to keep the existing number format |
| `NumberDecimalPlaces` | Decimals | `InArgument<int>` | `2` (when not set) | Number of decimal places (Number, Currency, and Percentage categories) |
| `NumberUseThousandSeparator` | Use 1000 Separator | `bool` | `False` | Use a thousands separator (Number and Currency categories) |
| `CurrencySymbol` | Symbol | `InArgument<string>` | `$` (when not set) | The currency symbol (Currency category) |
| `CurrencySetAtTheEnd` | Set at the end | `bool` | `False` | Place the currency symbol after the value instead of before it |
| `DateType` | Date format | `InArgument<DateType>` | `Short` (when not set) | The date format to use (Date category) |
| `TimeType` | Time format | `InArgument<TimeType>` | `HoursMinutes` (when not set) | The time format to use (Time category) |
| `TimeIsAMPM` | AM/PM | `bool` | `False` | Use 12-hour AM/PM format instead of 24-hour |
| `CustomFormat` | Custom format | `InArgument<string>` | | A custom Excel number format string (Custom category) |
| `HorizontalAlignment` | Horizontal | `InArgument<ExcelRangeHorizontalAlignment>` | | Horizontal alignment for cell content. Leave unset to keep the existing alignment |
| `VerticalAlignment` | Vertical | `InArgument<ExcelRangeVerticalAlignment>` | | Vertical alignment for cell content. Leave unset to keep the existing alignment |
| `WrapText` | Wrap text | `bool` | `False` | Wrap text within the cell. Applied when it is on, or when the horizontal or vertical alignment is set (so it can then turn the wrapping off); otherwise the existing wrapping is kept |
| `FontFamily` | Font | `InArgument<string>` | | The font family name. Leave unset to keep the existing font |
| `FontSize` | Font size | `InArgument<double>` | | The font size in points. Leave unset to keep the existing size |
| `FontStyle` | Font style | `InArgument<ExcelRangeFontStyle>` | | The font style: `Regular`, `Italic`, `Bold` or `BoldItalic`. Leave unset to keep the existing style |
| `FontColor` | Font color | `InArgument<string>` | | The font color: a color name (for example `Red`) or a hex color (`#RRGGBB`, or `#AARRGGBB`). Leave unset to keep the existing color |
| `FontUnderlineStyle` | Underline style | `InArgument<ExcelRangeUnderlineStyle>` | | The underline style of the font. Leave unset to keep the existing style (the designer hints None, the value the Windows activity starts with; choose None to remove an underline) |
| `FontFillColor` | Fill color | `InArgument<string>` | | The fill (background) color, as `FontColor`. `Transparent` removes the fill. Leave unset to keep the existing fill |
| `Password` | Password | `InArgument<string>` | | The password of the workbook, if necessary |
| `Workbook` | Workbook | `InArgument<Workbook>` | | Existing workbook object to use instead of a file path |

### Enum Reference

**`CellFormat`**: `General`, `Number`, `Date`, `Time`, `Percentage`, `Currency`, `Text`, `Custom`
**`DateType`**: `Long`, `Short`
**`TimeType`**: `HoursMinutes`, `HoursMinutesSeconds`
**`ExcelRangeHorizontalAlignment`**: `xlHAlignGeneral`, `xlHAlignLeft`, `xlHAlignCenter`, `xlHAlignRight`, `xlHAlignFill`, `xlHAlignJustify`, `xlHAlignCenterAcrossSelection`, `xlHAlignDistributed` (shared with the Windows `FormatRangeX` activity)
**`ExcelRangeFontStyle`**: `Regular`, `Italic`, `Bold`, `BoldItalic` (shared with the Windows `FormatRangeX` activity)
**`ExcelRangeUnderlineStyle`**: `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting` (shared with the Windows `FormatRangeX` activity)
**`ExcelRangeVerticalAlignment`**: `xlVAlignTop`, `xlVAlignCenter`, `xlVAlignBottom`, `xlVAlignJustify`, `xlVAlignDistributed` (shared with the Windows `FormatRangeX` activity)

## Designer

Like the Windows `FormatRangeX` activity, the designer shows one section of the settings at a time. The **Format** selector (Data Type, Alignment or Font) chooses the section; it is only used by the designer and is not saved with the workflow. The values of the sections that are not shown are kept and still applied.

- **Data Type** shows `Category`, and, once a category is chosen, only the fields of that category. While `Category` is unset nothing else applies, so only `Category` is shown. When `Category` is bound to a variable or an expression, every category field is shown.
- **Alignment** shows `HorizontalAlignment`, `VerticalAlignment` and `WrapText`.
- **Font** shows `FontFamily`, `FontStyle`, `FontSize`, `FontColor`, `FontUnderlineStyle` and `FontFillColor`, the same fields as the Windows activity.

## Valid Configurations

`Category`, when set, selects which number-format sub-properties apply:

- **`General`**: no sub-properties used, format is `General`.
- **`Number`**: `NumberDecimalPlaces`, `NumberUseThousandSeparator`.
- **`Currency`**: `NumberDecimalPlaces`, `NumberUseThousandSeparator`, `CurrencySymbol`, `CurrencySetAtTheEnd`.
- **`Percentage`**: `NumberDecimalPlaces`.
- **`Date`**: `DateType`.
- **`Time`**: `TimeType`, `TimeIsAMPM`.
- **`Text`**: no sub-properties, format is `@`.
- **`Custom`**: `CustomFormat` is used verbatim as the Excel number-format string.

`Category`, alignment (`HorizontalAlignment`, `VerticalAlignment`, `WrapText`) and font (`FontFamily`, `FontStyle`, `FontSize`, `FontColor`, `FontUnderlineStyle`, `FontFillColor`) properties are all independent of each other — any subset may be set. A property left unset (`Nothing`/`null`) leaves that aspect of the existing formatting untouched — including `Category`, so e.g. setting only `Bold` does not reset an existing number format to `General`.

## XAML Example

```xml
<excel:FormatRange
  DisplayName="Format Range Workbook"
  WorkbookPath="[&quot;report.xlsx&quot;]"
  SheetName="[&quot;Sheet1&quot;]"
  Range="[&quot;B2:B20&quot;]"
  Category="[CellFormat.Currency]"
  NumberDecimalPlaces="[2]"
  NumberUseThousandSeparator="True"
  CurrencySymbol="[&quot;$&quot;]"
  FontStyle="[ExcelRangeFontStyle.Bold]" />
```

## Notes

- None of the arguments is nullable. `WrapText` is a plain on/off option, like on the Windows activity, so it cannot be left unset: it is applied only when it is on or when the alignment is configured.
- None of the other arguments is nullable. An argument that has no value is not set, and leaves that aspect of the existing formatting untouched; an argument that has a value is applied, `False` and `0` included.

- Uses the file-based workbook engine (not Excel Interop). Works cross-platform for `.xlsx` files; throws `NotSupportedException` for `.xls` files.
- `WorkbookPath` and `WorkbookPathResource` are mutually exclusive — set one or the other.
- A whole-column or whole-row `Range` (e.g. `A:A`) applies a single column/row-level style (matching Excel's own behavior — it also applies to cells filled in later) instead of styling every one of the sheet's 1,048,576 rows/16,384 columns individually.
- Throws `ArgumentException` if `Range` is null/empty.
- `FontColor` and `FontFillColor` accept a color name (case-insensitive, any `System.Drawing` known color such as `Red` or `LightGray`) or a hex color; anything else throws `ArgumentException` before any option is applied. There is no color picker yet: the colors are typed.
- `Range` must refer to the current sheet (`SheetName`) — a sheet-qualified address (e.g. `Sheet2!A1:A10`) is not supported and throws `ArgumentException`.
- `CurrencySymbol` cannot contain a quote character — throws `ArgumentException` rather than being written to the file, since there's no way to represent it safely (doubling it produces two adjacent quoted sections instead of an escaped quote, silently dropping the character; escaping it with a backslash instead hangs the file the next time it's opened when the symbol comes first, the default layout, for the same reason as the next point). `CustomFormat` throws `ArgumentException` instead of being written to the file when, before the first number, date, or time placeholder character, the double quotes (escaped or not) don't occur in pairs, i.e. an opening quote has no later quote to close it — e.g. `"a"\"0.00`, `\""b"0.00`, or a bare `\"0.00` (rejected even though the backslash would read as an escape in Excel's own terms, since a double quote always counts toward this check whether or not it's backslash-escaped). Two forms are accepted: in `0.00\"` the escaped quote comes after the first placeholder, and in `"a\"0.00` the backslash is literal text inside the quoted section, so the quote right after it closes that section.
