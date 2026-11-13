"""Clean an Excel order sheet and create a summary workbook."""
import argparse
import json
from collections import defaultdict
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

REQUIRED_HEADERS = ["order_id", "product", "quantity", "unit_price", "status"]


def normalize_header(value):
    return "_".join(str(value or "").strip().lower().split())


def fit_columns(sheet):
    """Set readable widths based on the text in each column."""
    for column_cells in sheet.columns:
        letter = get_column_letter(column_cells[0].column)
        max_length = max(len(str(cell.value or "")) for cell in column_cells)
        sheet.column_dimensions[letter].width = min(max(max_length + 2, 12), 32)


def main():
    parser = argparse.ArgumentParser(description="Clean an Orders sheet and build a product summary.")
    parser.add_argument("input", nargs="?", default="messy_orders.xlsx", help="Input workbook")
    parser.add_argument("--output", default="orders_report.xlsx", help="Output workbook")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {input_path}. Run: python generate_data.py")

    source_book = load_workbook(input_path, data_only=True)
    source_sheet = source_book["Orders"] if "Orders" in source_book.sheetnames else source_book.active
    raw_headers = [cell.value for cell in source_sheet[1]]
    headers = [normalize_header(value) for value in raw_headers]
    missing = [header for header in REQUIRED_HEADERS if header not in headers]
    if missing:
        raise SystemExit(f"Missing required columns: {', '.join(missing)}")

    index = {name: headers.index(name) for name in REQUIRED_HEADERS}
    seen_ids = set()
    clean_rows = []
    blank_rows = 0
    duplicate_rows = 0
    invalid_rows = 0
    revenue_by_product = defaultdict(float)
    units_by_product = defaultdict(int)

    for values in source_sheet.iter_rows(min_row=2, values_only=True):
        if not any(value not in (None, "") for value in values):
            blank_rows += 1
            continue

        order_id = str(values[index["order_id"]] or "").strip()
        product = str(values[index["product"]] or "").strip()
        status = str(values[index["status"]] or "").strip().lower()
        quantity_value = values[index["quantity"]]
        price_value = values[index["unit_price"]]

        if not order_id or not product or quantity_value is None or price_value is None:
            invalid_rows += 1
            continue
        if order_id in seen_ids:
            duplicate_rows += 1
            continue
        try:
            quantity = int(quantity_value)
            unit_price = float(price_value)
        except (TypeError, ValueError):
            invalid_rows += 1
            continue
        if quantity <= 0 or unit_price < 0:
            invalid_rows += 1
            continue

        seen_ids.add(order_id)
        clean_rows.append([order_id, product, quantity, round(unit_price, 2), status])
        if status != "cancelled":
            units_by_product[product] += quantity
            revenue_by_product[product] += quantity * unit_price

    output_book = Workbook()
    clean_sheet = output_book.active
    clean_sheet.title = "Cleaned Orders"
    clean_sheet.append(["order_id", "product", "quantity", "unit_price", "status"])
    for row in clean_rows:
        clean_sheet.append(row)

    summary_sheet = output_book.create_sheet("Product Summary")
    summary_sheet.append(["product", "units_sold", "revenue"])
    for product in sorted(units_by_product):
        summary_sheet.append([product, units_by_product[product], round(revenue_by_product[product], 2)])

    report_sheet = output_book.create_sheet("Run Report")
    report_sheet.append(["metric", "value"])
    report_sheet.append(["input_data_rows", source_sheet.max_row - 1])
    report_sheet.append(["clean_output_rows", len(clean_rows)])
    report_sheet.append(["blank_rows_removed", blank_rows])
    report_sheet.append(["duplicate_rows_removed", duplicate_rows])
    report_sheet.append(["invalid_or_incomplete_rows_skipped", invalid_rows])

    for sheet in output_book.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="244062")
            cell.alignment = Alignment(horizontal="center")
        fit_columns(sheet)

    output_path = Path(args.output)
    output_book.save(output_path)

    report = {
        "input_file": str(input_path),
        "output_file": str(output_path),
        "input_data_rows": source_sheet.max_row - 1,
        "clean_output_rows": len(clean_rows),
        "blank_rows_removed": blank_rows,
        "duplicate_rows_removed": duplicate_rows,
        "invalid_or_incomplete_rows_skipped": invalid_rows,
    }
    report_path = output_path.with_suffix(".report.json")
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")

    print(json.dumps(report, indent=2))
    print(f"Saved workbook: {output_path}")
    print(f"Saved report: {report_path}")


if __name__ == "__main__":
    main()
