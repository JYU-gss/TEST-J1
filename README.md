# Quick Option Import/Export (recreated)

A small WinForms utility for your QA team to edit TA option JSON files.

## What it does
- Import a JSON file (array of option objects).
- Shows key fields in a grid:
  - Option Number, Sequence
  - Text / Time / Numeric / Long / Date / Boolean / ASCII
- Edit cells inline.
- Add new rows (use the bottom blank row).
- Delete rows (right-click → Delete Row).
- Save or Save As to export JSON.

## Build / Run (Windows)
1. Install the **.NET 8 SDK**.
2. Open `QuickOptionImportExport.sln` in Visual Studio 2022.
3. Build and run.

## Notes
- When you import an existing file, the app uses the first row as a template for new rows, so exports stay very close to the original format.
- Date format: `yyyy-MM-dd`
- Time format: `HH:mm` or `HH:mm:ss`
