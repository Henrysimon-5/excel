# Excel Sheet

A lightweight, browser-based spreadsheet application — no server, no build step, just open `index.html`.

## Features

| Category | Details |
|---|---|
| **Grid** | 50 rows × 26 columns (A – Z) |
| **Editing** | Double-click or press F2 / any printable key to edit a cell inline |
| **Formulas** | `=SUM`, `=AVERAGE`, `=MIN`, `=MAX`, `=COUNT`, `=IF`, `=ROUND`, `=ABS`, arithmetic expressions, cell references (e.g. `=A1+B2`), ranges (e.g. `=SUM(A1:A10)`) |
| **Formatting** | Bold, Italic, Underline, Strikethrough, Font size, Text colour, Fill colour, Align (L / C / R) |
| **Keyboard** | Arrow keys to navigate, Tab/Shift-Tab, Enter to confirm, Esc to cancel, Delete/Backspace to clear |
| **Clipboard** | Ctrl+C copy, Ctrl+X cut, Ctrl+V paste |
| **Undo / Redo** | Ctrl+Z / Ctrl+Y (100-step history) |
| **Context menu** | Right-click any cell for copy, cut, paste, clear, insert row, delete row |
| **Column resize** | Drag the right edge of any column header |
| **Export** | Click the CSV button to download the sheet as a `.csv` file |

## Usage

Open `index.html` directly in any modern browser (Chrome, Firefox, Edge, Safari).

```
open index.html   # macOS
start index.html  # Windows
xdg-open index.html  # Linux
```

No installation or internet connection required.
