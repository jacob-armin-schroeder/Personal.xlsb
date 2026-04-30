# Excel Personal Macro Package

A collection of Excel VBA macros designed to speed up common formatting and navigation tasks. All macros are intended for installation in Excel's **Personal Macro Workbook** (`PERSONAL.XLSB`), making them available across all workbooks.

---

## Modules

| File | Description |
|---|---|
| `PERSONAL_AutoFill.bas` | Fills a cell's content and formatting downward using an adjacent column as a length guide |
| `PERSONAL_CellFormatting.bas` | Cycles borders and fill colors on selected cells |
| `PERSONAL_FindChanges.bas` | Navigates a column by jumping to the next or previous value change |
| `PERSONAL_NumberFormats.bas` | Cycles number formats (decimal, percentage, currency, date/time) through common variants |
| `PERSONAL_RowColumnSize.bas` | Adjusts row height and column width, including autofit |

---

## Installation

1. Open Excel and press **Alt+F11** to open the Visual Basic Editor.
2. In the Project Explorer, expand **VBAProject (PERSONAL.XLSB)**.  
   *(If PERSONAL.XLSB does not exist, record any macro with "Personal Macro Workbook" selected as the store location, then delete the recorded macro — this creates the file.)*
3. Right-click **Modules** → **Import File**, and select each `.bas` file.
4. Save and close the VBE. Restart Excel if prompted.

### Assigning Keyboard Shortcuts

Each macro includes a recommended shortcut in its header comment. To assign them:

1. Go to **Developer** → **Macros** (or press Alt+F8).
2. Select a macro from the list and click **Options**.
3. Enter the shortcut key as noted in the table below.

> **Note:** Shortcuts listed as `Ctrl+Shift+Letter` require entering the uppercase letter in the shortcut field.

---

## Macro Reference

### AutoFill — `PERSONAL_AutoFill.bas`

**Shortcut:** `Ctrl+Shift+D`

Fills the active cell's content and formatting downward to the last continuous row used in an adjacent column.

- Prefers the **left** column as the length reference; falls back to the right.
- Column A always uses the right column.
- Stops autofilling at blank rows.
- Continues autofilling past `#N/A`, `#REF!` and other errors.

---

### Cell Formatting — `PERSONAL_CellFormatting.bas`

**Shortcuts:**

| Macro | Shortcut | Behavior |
|---|---|---|
| `Border_Table_Heading` | `Ctrl+H` | Applies bold, center-across-selection, wrap text, inside vertical lines, and a thin outer border |
| `VerticalLines` | `Ctrl+E` | Cycles inside vertical borders: thin → medium → none |
| `HorizontalLines` | `Ctrl+Shift+E` | Cycles inside horizontal borders: hairline → thin → medium → none |
| `Border_Outline` | `Ctrl+O` | Cycles outer border: thin → medium → none |
| `Clear_Formatting` | `Ctrl+N` | Removes all borders and fill; resets font color to automatic |
| `FillBright` | `Ctrl+Shift+B` | Cycles background through light colors: gray → blue → orange → green → white → none |
| `FillDark` | `Ctrl+Shift+D` | Cycles background through dark colors (with white font): gray → blue → orange → green → black → none |

> **Note:** These macros cycle independently based on the active cell's current formatting. Formatting outside a given cycle is treated as "no formatting" so that the cycle starts at the beginning.

---

### Find Changes — `PERSONAL_FindChanges.bas`

**Shortcuts:** `Ctrl+M` (next) / `Ctrl+Shift+M` (previous)

Navigates the active column by jumping to the next (or previous) cell whose value differs from the current cell. Useful for stepping through columns where values repeat in blocks — status codes, categories, flags, etc.

- Stays within the sheet's used range.
- If no different value is found, lands on the last cell evaluated.
- Briefly highlights the destination cell in yellow, then restores the original fill.
- Handles error values as distinct, comparable strings.

---

### Number Formats — `PERSONAL_NumberFormats.bas`

**Shortcuts:**

| Macro | Shortcut | Cycle |
|---|---|---|
| `NumberFormatDecimal` | `Ctrl+Shift+A` | `#,##0` → `.0` → `.00` → `.000` → `.0000` → `.00000` → `.000000` → (repeat) |
| `NumberFormatPercentage` | `Ctrl+Shift+P` | `#,##0%` → `.0%` → `.00%` → `.000%` → (repeat) |
| `NumberFormatCurrency` | `Ctrl+Shift+C` | `$#,##0` → `$#,##0.00` → red-negative variants → accounting-style variants → (repeat) |
| `NumberFormatDateTime` | `Ctrl+Shift+T` | Short date → long date → zero-padded date → time → datetime combinations → ISO format → (repeat); autofits column |

All format macros (except DateTime) apply right alignment and turn off wrap text.

---

### Row and Column Size — `PERSONAL_RowColumnSize.bas`

**Shortcuts:**

| Macro | Shortcut | Behavior |
|---|---|---|
| `Autofit` | `Ctrl+Shift+W` | Autofits both column width and row height for the selection |
| `ColumnWidthIncrease` | `Ctrl+Q` | Increases column width by 1 unit |
| `ColumnWidthDecrease` | `Ctrl+Shift+Q` | Decreases column width by 1 unit (minimum: 1) |
| `RowHeightIncrease` | `Ctrl+R` | Increases row height by 5 units |
| `RowHeightDecrease` | `Ctrl+Shift+R` | Decreases row height by 5 units (minimum: 5) |

---

## Compatibility

Tested in Excel for Windows. Macros use standard Excel object model calls and should be compatible with Excel 2016 and later. Not tested in Excel for Mac.

---

## Notes

- All modules use `Option Explicit`, requiring all variables to be declared before use.
- Shortcut conflicts with existing Excel defaults are possible — verify against your Excel version before assigning.
- These macros operate on the active cell or selection and do not modify any other workbook state.