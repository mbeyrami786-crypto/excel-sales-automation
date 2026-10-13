# Excel Sales Automation

An automated Excel reporting system built with Python, Pandas and OpenPyXL,
with an optional Streamlit front-end.

## Features

- Reads sales data from an Excel file
- Validates columns and cleans the data (trims text, coerces numbers,
  drops rows with missing or non-positive `Quantity`/`Price`)
- Automatically calculates revenue
- Total revenue, total quantity, average unit price, weighted price,
  transaction count
- Groups sales by product and by region
- Top 5 products and Top 5 regions by revenue
- Creates a styled Excel **Dashboard** with KPI cards
- Creates bar and pie **charts** automatically
- Streamlit web UI that downloads the very same report

## Requirements

- Python 3.10+
- `pip install -r requirements.txt`

## Project Structure

```text
excel_automation/
│
├── core.py               # shared logic: validation, metrics, report builder
├── main3.py              # CLI entry point
├── app.py                # Streamlit web UI
├── generate_data.py      # writes sample data to input/sales.xlsx
├── requirements.txt
├── README.md
│
├── input/
│   └── sales.xlsx        # sample data (tracked; replace with your own)
│
├── output/
│   └── premium_automated_report.xlsx   # generated report (not tracked)
│
└── tests/                # unit + UI tests
```

Generated reports (`output/*.xlsx`) and old revisions kept under `archive/`
are excluded from version control by `.gitignore`.

## Usage

```bash
# 1. install dependencies
pip install -r requirements.txt

# 2. (optional) create sample data
python generate_data.py

# 3. build the report
python main3.py

# 4. or use the web UI
streamlit run app.py
```

### CLI options

```text
python main3.py --input input/sales.xlsx --output output/report.xlsx --currency $
```

| Option       | Default                        | Description                    |
| ------------ | ------------------------------ | ------------------------------ |
| `--input`    | `input/sales.xlsx`             | source workbook                |
| `--output`   | `output/premium_automated_report.xlsx` | destination workbook  |
| `--currency` | `$`                            | currency symbol in the report  |

Paths are resolved from the project folder, so the scripts work from any
current working directory.

## Tests

```bash
python -m unittest discover -s tests -v
```

The suite covers the cleaning/validation rules, the metrics, the generated
workbook (sheets, charts, number formats) and the Streamlit UI.

## Input format

The input workbook must contain these columns:

| Column    | Type   | Rules                          |
| --------- | ------ | ------------------------------ |
| Product   | text   | required, blank rows discarded |
| Region    | text   | required, blank rows discarded |
| Quantity  | number | required, must be > 0          |
| Price     | number | required, must be > 0          |

Invalid rows are dropped and reported (`Rows Discarded` in the web UI).
If a required column is missing, a clear `DataValidationError` is raised
instead of a cryptic traceback.

## Output

`output/premium_automated_report.xlsx` contains:

1. **Dashboard** – KPI cards, top product/region, bar + pie charts
2. **Business Insights** – metric table plus the Top 5 products/regions
3. **Sales Data** – cleaned rows with the calculated `Revenue`
4. **By Product** – revenue per product (descending)
5. **By Region** – revenue per region (descending)
