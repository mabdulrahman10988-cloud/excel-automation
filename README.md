# Excel Automation

A Python tool that processes an Excel `Orders` worksheet, trims text, removes blank and duplicate orders, skips incomplete or invalid rows, and creates a formatted workbook with cleaned orders, product totals, and a run report.

**Portfolio status:** Independent demonstration using synthetic data. It is not a commissioned client project.

## Requirements

- Python 3.10 or newer
- `openpyxl` (installed from `requirements.txt`)

## Install and run

```bash
python -m pip install -r requirements.txt
python generate_data.py
python excel_automation.py
```

Expected files:

- `messy_orders.xlsx` — generated sample workbook
- `orders_report.xlsx` — output workbook with three sheets
- `orders_report.report.json` — machine-readable run summary

## Process another workbook

The input workbook must contain an `Orders` sheet (or use the first sheet) with these columns:

- `Order ID`
- `Product`
- `Quantity`
- `Unit Price`
- `Status`

Run:

```bash
python excel_automation.py path/to/orders.xlsx --output cleaned_report.xlsx
```

## Output sheets

1. **Cleaned Orders** — normalized rows.
2. **Product Summary** — units and revenue by product; cancelled orders are excluded.
3. **Run Report** — input, output, blank, duplicate and invalid row counts.

## Limits

- Rows missing order ID, product, quantity or unit price are skipped and counted.
- Duplicate detection uses the order ID.
- This demonstration does not connect to an online store or accounting system.
- Check business rules before using the output for accounting or inventory decisions.

## Reproducibility

The generator uses a fixed random seed, so it creates the same sample data each time.

## License

MIT. See `LICENSE` if you add one before public distribution.
