# Excel to Memo

`excel-to-memo` is a small Python CLI package that converts a monitoring-detail Excel or CSV file into a memo-style review output.

It is designed around the structure shown in your screenshots:

- Source A: row-level monitoring details with columns like `Section`, `Model Parameter`, `Portfolio`, `Status`, `ECL impact`, and `Comments`
- Target B: a memo/report summary with review metadata, assessment summary, and section-level comments

## What it generates

- `memo_review.docx`
- `memo_review.md`
- `normalized_rows.csv`

## Comment generation rules

The memo body now follows these rules:

- If a grouped portfolio/section has no `Fail` rows, the tool copies the Excel `Comments` content directly, while removing exact duplicates from repeated rows.
- If a grouped portfolio/section has one or more `Fail` rows, the tool ignores the non-fail narrative in that group and builds the bullet from the failed rows only.
- For failed rows, the tool prefers the row's own Excel `Comments` text. If that cell is blank, it falls back to a sentence built from `Model Parameter`, `Over/under estimate`, `ECL impact`, and `Materiality`.

This means the Word file is treated as a layout template, not as a content source.

## Quick start

### Recommended: run directly without installation

First install the runtime dependencies only:

```bash
pip install -r requirements-py37.txt
```

From the repo root:

```bash
python run_excel_to_memo.py input.xlsx --config config.json --output-dir out
```

If your sheet already uses the expected column names, you can also omit the config:

```bash
python run_excel_to_memo.py input.xlsx --output-dir out
```

### Alternative: run the module from source

```bash
pip install -r requirements-py37.txt
PYTHONPATH=src python -m excel_to_memo input.xlsx --config config.json --output-dir out
```

`config.json` is included at the repo root as a ready-to-run starter config.

### Optional: install the package

Only do this if your environment can install build dependencies.

```bash
python -m pip install .
```

Then run:

```bash
excel-to-memo input.xlsx --config config.json --output-dir out
```

If the console script is not on your `PATH`, use the module form instead:

```bash
python -m excel_to_memo input.xlsx --config config.json --output-dir out
```

## Windows

The package is designed to work on Windows as a normal Python CLI.

Command Prompt:

```bat
py -m pip install .
py -m excel_to_memo input.xlsx --config config.json --output-dir out
```

PowerShell:

```powershell
python -m pip install .
python -m excel_to_memo .\\input.xlsx --config .\\config.json --output-dir .\\out
```

Notes for Windows:

- `.xlsx` reading uses `openpyxl`
- legacy `.xls` reading uses `xlrd`
- output paths can be relative or absolute
- the generated files are standard `.docx`, `.md`, and `.csv`
- Python `3.7+` is supported

## Python version support

- Python `3.7` to current versions are supported
- On Python `3.7`, the package installs a `pandas 1.x` compatible dependency set
- On Python `3.8+`, the package installs the newer `pandas 2.x` path

## Local development

Without installing the package, you can run it from source:

```bash
pip install -r requirements-py37.txt
PYTHONPATH=src python -m excel_to_memo input.xlsx --config config.json --output-dir out
```

Or use the direct runner at the repo root:

```bash
pip install -r requirements-py37.txt
python run_excel_to_memo.py input.xlsx --config config.json --output-dir out
```

## Expected input columns

By default the tool looks for the following business columns, with flexible header matching:

- `Section`
- `Sub-section`
- `Model name`
- `Model Parameter`
- `Portfolio`
- `Status`
- `Materiality`
- `Over/under estimate`
- `ECL impact`
- `Comments`

You can override the mapping in the config file.

## Config

The config file lets you control:

- report title and metadata
- source sheet name
- grouping key for memo bullets
- exact column mapping if your workbook headers differ
- default checklist items in the memo

See `examples/sample_config.json`.

## Notes

- Input can be `.xlsx`, `.xls`, `.csv`, or `.tsv`
- Direct-run mode still needs the runtime dependencies from `requirements-py37.txt`
- If a row already has narrative text in `Comments`, that text is preferred in the generated memo
- If comments are blank, the tool falls back to structured sentences built from status, parameter, estimate direction, and ECL impact
- The sample workbook is only illustrative. Final section names and comment text always come from the source Excel file you run through the tool.
