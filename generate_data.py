"""Create a deliberately messy Excel workbook for the automation demo."""
from pathlib import Path
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

random.seed(31)
OUTPUT = Path("messy_orders.xlsx")
products = ["Notebook", "Desk Lamp", "Water Bottle", "Backpack", "USB Hub"]
workbook = Workbook()
sheet = workbook.active
sheet.title = "Orders"
sheet.append([" Order ID ", "Product", "Quantity", "Unit Price", "Status"])

rows = []
for number in range(1, 61):
    row = [
        f" O-{number:04d} ",
        f" {random.choice(products)} ",
        random.choice([1, 2, 3, 4, 5, None]),
        random.choice([9.99, 14.50, 22.00, 35.75, None]),
        random.choice([" paid ", "pending", " Paid", "cancelled "]),
    ]
    if number % 14 == 0:
        row[0] = " "
    rows.append(row)

# Add a duplicate order and one completely blank line.
rows.append(rows[5].copy())
rows.insert(12, [None, None, None, None, None])
for row in rows:
    sheet.append(row)

for cell in sheet[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="244062")
sheet.freeze_panes = "A2"
workbook.save(OUTPUT)
print(f"Created {OUTPUT} with {len(rows)} data rows plus a header.")
