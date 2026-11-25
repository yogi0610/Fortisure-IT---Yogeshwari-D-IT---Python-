from pathlib import Path
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
project_folder = Path(__file__).resolve().parent.parent
input_file = project_folder / "cleaned_data" / "Quick_Bite_Combined_Data.xlsx"
output_folder = project_folder / "analysis"
output_folder.mkdir(exist_ok=True)

if not input_file.exists():
    raise FileNotFoundError(
        "Quick_Bite_Combined_Data.xlsx was not found. Run cleaning.py first."
    )

data = pd.read_excel(input_file)
required_columns = [
    "order_id",
    "order_day",
    "cuisine",
    "area_name",
    "city",
    "vehicle_type",
    "order_amount",
    "order_status",
    "delivery_minutes",
    "rating",
]

missing_columns = [col for col in required_columns if col not in data.columns]
if missing_columns:
    raise ValueError(f"These columns are missing from the combined file: {missing_columns}")

data["order_amount"] = pd.to_numeric(data["order_amount"], errors="coerce")
data["delivery_minutes"] = pd.to_numeric(
    data["delivery_minutes"], errors="coerce"
)
data["rating"] = pd.to_numeric(data["rating"], errors="coerce")
data["order_status"] = (
    data["order_status"].astype("string").str.strip().str.lower()
)
data["is_late"] = data["delivery_minutes"] > 45
area_performance = (
    data.groupby(["area_name", "city"], dropna=False)
    .agg(
        total_orders=("order_id", "nunique"),
        revenue=("order_amount", "sum"),
        average_order_value=("order_amount", "mean"),
        average_delivery_time=("delivery_minutes", "mean"),
        late_orders=("is_late", "sum"),
        orders_with_delivery_time=("delivery_minutes", "count"),
        average_rating=("rating", "mean"),
    )
    .reset_index()
)

area_performance["late_rate_percent"] = np.where(
    area_performance["orders_with_delivery_time"] > 0,
    area_performance["late_orders"]
    / area_performance["orders_with_delivery_time"]
    * 100,
    np.nan,
)

area_performance["on_time_rate_percent"] = np.where(
    area_performance["orders_with_delivery_time"] > 0,
    (
        area_performance["orders_with_delivery_time"]
        - area_performance["late_orders"]
    )
    / area_performance["orders_with_delivery_time"]
    * 100,
    np.nan,
)

area_performance = area_performance.drop(
    columns=["orders_with_delivery_time"]
)
area_performance = area_performance.sort_values(
    "revenue", ascending=False
).reset_index(drop=True)

area_performance.insert(
    0, "revenue_rank", range(1, len(area_performance) + 1)
)

cuisine_orders = (
    data.dropna(subset=["order_day", "cuisine"])
    .groupby(["order_day", "cuisine"])["order_id"]
    .nunique()
    .reset_index(name="order_count")
)
daily_order_totals = (
    data.dropna(subset=["order_day"])
    .groupby("order_day")["order_id"]
    .nunique()
)

cuisine_orders["daily_total_orders"] = (
    cuisine_orders["order_day"].map(daily_order_totals)
)
cuisine_orders["daily_share_percent"] = (
    cuisine_orders["order_count"]
    / cuisine_orders["daily_total_orders"]
    * 100
)

top_cuisine_by_day = (
    cuisine_orders.sort_values(
        ["order_day", "order_count", "cuisine"],
        ascending=[True, False, True],
    )
    .drop_duplicates(subset="order_day", keep="first")
    .copy()
)

day_order = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
]

top_cuisine_by_day["order_day"] = pd.Categorical(
    top_cuisine_by_day["order_day"],
    categories=day_order,
    ordered=True,
)
top_cuisine_by_day = top_cuisine_by_day.sort_values(
    "order_day"
).reset_index(drop=True)

top_cuisine_by_day = top_cuisine_by_day[
    [
        "order_day",
        "cuisine",
        "order_count",
        "daily_total_orders",
        "daily_share_percent",
    ]
]
delivered_orders = data[
    data["order_status"].eq("delivered")
].copy()

vehicle_performance = (
    delivered_orders.dropna(subset=["vehicle_type"])
    .groupby("vehicle_type")
    .agg(
        delivered_orders=("order_id", "nunique"),
        average_delivery_time=("delivery_minutes", "mean"),
        orders_with_delivery_time=("delivery_minutes", "count"),
        late_orders=("is_late", "sum"),
    )
    .reset_index()
)

vehicle_performance["on_time_orders"] = (
    vehicle_performance["orders_with_delivery_time"]
    - vehicle_performance["late_orders"]
)

vehicle_performance["late_rate_percent"] = np.where(
    vehicle_performance["orders_with_delivery_time"] > 0,
    vehicle_performance["late_orders"]
    / vehicle_performance["orders_with_delivery_time"]
    * 100,
    np.nan,
)

vehicle_performance["on_time_rate_percent"] = np.where(
    vehicle_performance["orders_with_delivery_time"] > 0,
    vehicle_performance["on_time_orders"]
    / vehicle_performance["orders_with_delivery_time"]
    * 100,
    np.nan,
)

vehicle_performance["orders_without_delivery_time"] = (
    vehicle_performance["delivered_orders"]
    - vehicle_performance["orders_with_delivery_time"]
)

vehicle_performance = vehicle_performance.drop(
    columns=["orders_with_delivery_time"]
)
vehicle_performance = vehicle_performance.sort_values(
    ["on_time_rate_percent", "average_delivery_time"],
    ascending=[False, True],
    na_position="last",
).reset_index(drop=True)


output_file = output_folder / "Quick_Bite_Analysis.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    area_performance.to_excel(
        writer, sheet_name="Area Performance", index=False
    )
    top_cuisine_by_day.to_excel(
        writer, sheet_name="Top Cuisine by Day", index=False
    )
    vehicle_performance.to_excel(
        writer, sheet_name="Vehicle Performance", index=False
    )
workbook = load_workbook(output_file)

for sheet in workbook.worksheets:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    # Format the header row
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            fill_type="solid", fgColor="1F4E78"
        )
        cell.alignment = Alignment(
            horizontal="center", vertical="center"
        )
    for column_cells in sheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        longest_value = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in column_cells
        )
        sheet.column_dimensions[column_letter].width = min(
            max(longest_value + 3, 15), 32
        )
    for column_cells in sheet.columns:
        heading = str(column_cells[0].value).lower()
        if "percent" in heading:
            for cell in column_cells[1:]:
                cell.number_format = "0.00"
    for column_cells in sheet.columns:
        heading = str(column_cells[0].value).lower()
        if heading in ["revenue", "average_order_value"]:
            for cell in column_cells[1:]:
                cell.number_format = '#,##0.00'
    for column_cells in sheet.columns:
        heading = str(column_cells[0].value).lower()
        if heading in ["average_delivery_time", "average_rating"]:
            for cell in column_cells[1:]:
                cell.number_format = "0.00"

workbook.save(output_file)

print("Area performance records:", len(area_performance))
print("Top cuisine days:", len(top_cuisine_by_day))
print("Vehicle types analysed:", len(vehicle_performance))
print("Analysis Excel file saved to:", output_file)

