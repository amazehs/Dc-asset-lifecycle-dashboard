# scripts

Run from the repository root after `pip install -r requirements.txt`:

| Script | What it does |
|---|---|
| generate_data.py | Creates the synthetic CSV data in `data/` |
| build_workbook.py | Builds the workbook in `output/` with formulas, tables, validation, formatting and charts |
| validate_workbook.py | Recomputes 81 figures with pandas and compares them with a calculated copy of the workbook |
| change_test.py | Applies four realistic edits to a copy of the workbook |
| verify_change_test.py | Checks that the dashboard moved exactly as expected after those edits |

```bash
python scripts/generate_data.py
python scripts/build_workbook.py
```
