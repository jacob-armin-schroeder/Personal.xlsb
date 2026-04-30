# Excel Personal Macro Package

A collection of Excel VBA macros designed to speed up common formatting and navigation tasks. All macros are intended for installation in Excel's **Personal Macro Workbook** (`PERSONAL.XLSB`), making them available across all workbooks.

See SHORTCUTS.md for a quick reference of all keyboard shortcuts.

---

## Modules

| File | Description |
|---|---|
| `PERSONAL_AutoFill.bas` | Fills a cell's content and formatting downward using an adjacent column as a length guide |
| `PERSONAL_CellFormatting.bas` | Cycles borders and fill colors on selected cells |
| `PERSONAL_FindChanges.bas` | Navigates a column by jumping to the next or previous value change |
| `PERSONAL_NumberFormats.bas` | Cycles number formats (decimal, percentage, currency, date/time) through common variants |
| `PERSONAL_RowColumnSize.bas` | Adjusts row height and column width, including autofit |
| `PERSONAL_Installer.bas` | One-time installer: downloads all modules from GitHub and assigns shortcuts |

---

## Installation

### Automatic (recommended)

1. Download `PERSONAL_Installer.bas` from this repository.
2. Open Excel and press **Alt+F11** to open the Visual Basic Editor.
3. In the Project Explorer, right-click any module under **VBAProject (PERSONAL.XLSB)** → **Import File** → select `PERSONAL_Installer.bas`.
*(If PERSONAL.XLSB does not exist, record any macro with "Personal Macro Workbook" selected as the store location, then delete the recorded macro — this creates the file.)*
4. Press **Alt+F8**, select InstallPackage, and click **Run**.
5. Follow the prompts. The installer will download all modules, offer to assign keyboard shortcuts, and confirm when complete.
6. After installation, delete the installer module: right-click **PERSONAL_Installer** in the Project Explorer → **Remove Module**.

> **Requirement:** The installer needs access to the VBA project object model. If it reports an access error, go to **File → Options → Trust Center → Trust Center Settings → Macro Settings** and check **Trust access to the VBA project object model**.

### Manual
 
1. Download each `.bas` file from this repository.
2. Open Excel and press **Alt+F11**.
3. In the Project Explorer, right-click **Modules** under **VBAProject (PERSONAL.XLSB)** → **Import File**, and select each `.bas` file.
4. Assign keyboard shortcuts manually: press **Alt+F8**, select a macro, click **Options**, and enter the shortcut key. See [SHORTCUTS.md](SHORTCUTS.md) for the full list.
> **Note:** `Ctrl+Shift+Letter` shortcuts require entering the uppercase letter in the shortcut field.
 
---
 
## Macro Reference

### AutoFill — `PERSONAL_AutoFill.bas`

**Shortcut:** `Ctrl+D`

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
| `Border_Table_Heading` | `Ctrl+Shift+H` | Applies bold, center-across-selection, wrap text, inside vertical lines, and a thin outer border |
| `VerticalLines` | `Ctrl+E` | Cycles inside vertical borders: thin → medium → none |
| `HorizontalLines` | `Ctrl+Shift+E` | Cycles inside horizontal borders: hairline → thin → medium → none |
| `Border_Outline` | `Ctrl+Shift+O` | Cycles outer border: thin → medium → none |
| `Clear_Formatting` | `Ctrl+Shift+N` | Removes all borders and fill; resets font color to automatic |
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

## Customizing
 
All modules are open for editing. The macros most likely to benefit from customization are the cycling ones:
 
- **`FillBright` / `FillDark`** — color arrays are defined at the top of each sub. Add, remove, or replace RGB values to change the cycle.
- **`NumberFormatCurrency`** — the format string array can be trimmed or extended with additional currency formats.
- **`NumberFormatDateTime`** — add or remove format strings to match the date conventions used in your work.
When editing, note that all modules use `Option Explicit` — any new variables must be declared with `Dim` before use.
 
---

## Compatibility

Tested in Excel for Windows (Excel 2016 and later). Not tested in Excel for Mac.

---

## Notes

- These macros operate on the active cell or selection and do not modify any other workbook state.
- The installer requires an internet connection to download files from GitHub. For offline installation, use the manual method.
- Shortcut assignments are stored in PERSONAL.XLSB and persist across Excel sessions.