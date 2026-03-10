# Cube Data Changer AIO Manual

All-in-one tool for generating and processing concrete and mortar cube test data, now with manual cell mapping for different office Excel formats.

## What This Version Adds

- Manual configuration for grade, date, weight, and strength cells
- Works with non-standard office templates (different rows/columns)
- Same workflow as AIO, but not locked to fixed cell positions

## Features

- Auto-detect grade/type from each sheet (using configurable grade cell)
- Generate test data for:
	- Concrete grades: M10, M15, M20, M25, M30, M35, M40, M45
	- Mortar types: 1:4, 1:6
- Apply calendar dates (7-day and 28-day)
- Legacy mode for existing grade Excel files
- Dark modern desktop UI built with CustomTkinter
- Cross-platform settings saved to local JSON

## Processing Modes

- Auto Detect + Generate + Date
- Auto Detect + Generate
- Date Only
- Files + Date (Legacy)
- Files Only (Legacy)

## Quick Start

```bash
pip install -r requirements.txt
python app.py
```

## Manual Cell Configuration

Click the `Cell Configuration` button in the app header and set your office format.

### Default Mapping

- Grade Cell: `B12`
- Casting Date Cell: `C17`
- 7-Day Date Cell: `C18`
- 28-Day Date Cell: `F18`
- Weight Row / Start Column / Count: `25 / C / 6`
- 7-Day Strength Row / Start Column / Count: `27 / C / 3`
- 28-Day Strength Row / Start Column / Count: `27 / F / 3`

### Validation Rules

- Cell refs must be like `B12`, `AA27`
- Column refs must be like `C`, `AF`
- Row and Count values must be positive numbers

## Project Files

- `app.py`: GUI application
- `processor.py`: template processing logic and manual cell map handling
- `generator.py`: test data generator
- `settings.py`: persistent settings storage
- `requirements.txt`: dependencies

## Output

The processed file is saved as:

`<OriginalTemplateName>_Processed.xlsx`

inside your selected output folder.

## Notes

- The app preserves your last-used paths, mode, and cell mapping.
- If your office format changes, update only the mapping, not the code.

## License

MIT License
