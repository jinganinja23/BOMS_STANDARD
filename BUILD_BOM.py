import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path

BOM_FOLDER = Path("BOMs")
OUTPUT_FILE = Path("BOM_Standardisation.xlsx")

# tries description first, falls back to designator prefix
def classify(description, designator):
    desc = str(description).lower()
    if "capacitor" in desc:return "Capacitor"
    if "resistor" in desc:return "Resistor"
    if "led" in desc:return "LED"
    if "diode" in desc:return "Diode"
    if "inductor" in desc:return "Inductor"
    if "transistor" in desc:return "Transistor"
    if "connector" in desc:return "Connector"
    if "oscillator" in desc or "crystal" in desc:return "Oscillator"
    if "switch" in desc or "button" in desc:return "Switch"

    des = str(designator).upper()
    if des.startswith("DS"):return "LED"
    if des.startswith("C"):return "Capacitor"
    if des.startswith("R"):return "Resistor"
    if des.startswith("D"):return "Diode"
    if des.startswith("L"):return "Inductor"
    if des.startswith("J"):return "Connector"
    if des.startswith("SW"):return "Switch"
    if des.startswith("Y") or des.startswith("X"): return "Oscillator"
    if des.startswith("Q"): return "Transistor"

    return "IC"

# using to test the scrip, in production, you can replace this with a loop that reads all CSVs in the folder
files = ["CMU.csv", "BMU.csv"]

frames = []
for file in files:
    csv_path = BOM_FOLDER / file
    df_temp = pd.read_csv(csv_path)
    df_temp = df_temp.drop(columns=df_temp.columns[0])
    df_temp = df_temp.drop(columns=["Supplier Unit Price 1", "Supplier Subtotal 1", "Revision ID","Revision Status", "Manufacturer 1", "Manufacturer Lifecycle 1",
                                    "Supplier 1", "Supplier Part Number 1"], errors="ignore")
    df_temp = df_temp.rename(columns={
        "Line #": "line",
        "Quantity": "qty",
        "Name": "name",
        "Description": "description",
        "Manufacturer Part Number 1": "part_number",
        "Designator": "designator",
    })
    df_temp["class"] = df_temp.apply(lambda row: classify(row["description"], row["designator"]), axis=1)
    df_temp["board"] = csv_path.stem
    frames.append(df_temp)

df = pd.concat(frames, ignore_index=True)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Raw Data"

headers = ["Board", "Class", "Name", "Description", "Qty", "Designator", "Part Number"]
for col, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(name="Arial", bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", start_color="1A2C5B", fgColor="1A2C5B")
    cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 28

for row_idx, (_, row) in enumerate(df.iterrows(), 2):
    ws.cell(row=row_idx, column=1, value=row["board"])
    ws.cell(row=row_idx, column=2, value=row["class"])
    ws.cell(row=row_idx, column=3, value=row["name"])
    ws.cell(row=row_idx, column=4, value=row["description"])
    ws.cell(row=row_idx, column=5, value=row["qty"])
    ws.cell(row=row_idx, column=6, value=row["designator"])
    ws.cell(row=row_idx, column=7, value=row["part_number"])

ws.auto_filter.ref = f"A1:G{len(df) + 1}"
ws.column_dimensions["A"].width = 16
ws.column_dimensions["B"].width = 13
ws.column_dimensions["C"].width = 20
ws.column_dimensions["D"].width = 48
ws.column_dimensions["E"].width = 8
ws.column_dimensions["F"].width = 30
ws.column_dimensions["G"].width = 20

wb.save(OUTPUT_FILE)
print(f"done, saved to {OUTPUT_FILE}")