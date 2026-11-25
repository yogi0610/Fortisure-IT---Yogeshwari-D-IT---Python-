import re
from datetime import datetime
from pathlib import Path
import numpy as np
import pandas as pd
project_folder = Path(__file__).resolve().parent.parent
input_file = project_folder / "data" / "QuickBite_Assessment.xlsx"
output_folder = project_folder / "cleaned_data"
output_folder.mkdir(exist_ok=True)

data = pd.read_excel(input_file, sheet_name=None)
areas = data["areas"]
restaurants = data["restaurants"]
customers = data["customers"]
partners = data["delivery_partners"]
orders = data["orders"]
feedback = data["feedback"]
def strip_text(df):
    for col in df.columns:
        df[col] = df[col].apply(
            lambda value: value.strip() if isinstance(value, str) else value
        )
    return df.replace("", np.nan)
def clean_id(df, columns):
    for col in columns:
        df[col] = df[col].apply(
            lambda value: value.strip().upper() if isinstance(value, str) else value
        )
    return df
def to_date(value):
    if pd.isna(value):
        return pd.NaT

    date_text = str(value).strip()
    date_formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M",
        "%d %b %Y %I:%M %p",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d-%b-%Y",
    ]

    for date_format in date_formats:
        try:
            return pd.Timestamp(datetime.strptime(date_text, date_format))
        except ValueError:
            pass

    return pd.NaT

def remove_duplicates(df, id_column):
    df = df.drop_duplicates()
    return df.drop_duplicates(subset=id_column, keep="first")

def drop_missing(df, columns):
    return df.dropna(subset=columns)

def clean_phone(value):
    if pd.isna(value):
        return np.nan

    digits = re.sub(r"\D", "", str(value))
    return digits[-10:] if len(digits) >= 10 else np.nan

def standard(series, name_map):
    return series.str.lower().map(name_map)

def show_result(name, before, after):
    print(f"{name}: {before} rows before cleaning, {len(after)} rows after cleaning")
city_names = {
    "blr": "Bengaluru",
    "bengaluru": "Bengaluru",
    "hyd": "Hyderabad",
    "hyderabad": "Hyderabad",
    "chennai": "Chennai",
    "madras": "Chennai",
    "mumbai": "Mumbai",
}

cuisine_names = {
    "chinese": "Chinese",
    "biryani": "Biryani",
    "biriyani": "Biryani",
    "burgers": "Burgers",
    "desserts": "Desserts",
    "dessert": "Desserts",
    "south indian": "South Indian",
    "south-indian": "South Indian",
    "north indian": "North Indian",
    "north-indian": "North Indian",
    "n. indian": "North Indian",
    "pizza": "Pizza",
    "italian - pizza": "Pizza",
}

vehicle_names = {
    "bicycle": "Bicycle",
    "cycle": "Bicycle",
    "bike": "Motorbike",
    "motorbike": "Motorbike",
    "motor bike": "Motorbike",
    "scooter": "Scooter",
    "scooty": "Scooter",
    "activa": "Scooter",
    "ev": "EV Scooter",
    "ev scooter": "EV Scooter",
    "e-scooter": "EV Scooter",
}

payment_names = {
    "upi": "UPI",
    "gpay": "UPI",
    "phonepe": "UPI",
    "card": "Card",
    "debit card": "Card",
    "credit card": "Card",
    "cash": "Cash",
    "cod": "Cash",
    "cash on delivery": "Cash",
    "wallet": "Wallet",
    "paytm wallet": "Wallet",
    "amazon pay": "Wallet",
}

status_names = {
    "delivered": "Delivered",
    "completed": "Delivered",
    "cancelled": "Cancelled",
    "canceled": "Cancelled",
}

before = len(areas)
areas = strip_text(areas)
areas = clean_id(areas, ["area_id"])
areas["city"] = standard(areas["city"], city_names)
areas = drop_missing(areas, ["area_id", "area_name", "city"])
areas = remove_duplicates(areas, "area_id")
show_result("Areas", before, areas)

before = len(restaurants)
restaurants = strip_text(restaurants)
restaurants = clean_id(restaurants, ["restaurant_id", "area_id"])
restaurants["restaurant_name"] = restaurants["restaurant_name"].str.title()
restaurants["cuisine"] = standard(restaurants["cuisine"], cuisine_names)
restaurants["onboard_date"] = restaurants["onboard_date"].apply(to_date)
restaurants["rating"] = pd.to_numeric(restaurants["rating"], errors="coerce")
restaurants = drop_missing(
    restaurants,
    ["restaurant_id", "restaurant_name", "area_id", "cuisine", "onboard_date"],
)
restaurants = remove_duplicates(restaurants, "restaurant_id")
show_result("Restaurants", before, restaurants)

before = len(customers)
customers = strip_text(customers)
customers = clean_id(customers, ["customer_id", "area_id"])
customers["customer_name"] = customers["customer_name"].str.title()
customers["phone"] = customers["phone"].apply(clean_phone)
customers["email"] = customers["email"].str.lower()
customers["signup_date"] = customers["signup_date"].apply(to_date)
customers = drop_missing(
    customers, ["customer_id", "customer_name", "phone", "signup_date"]
)
customers = remove_duplicates(customers, "customer_id")
show_result("Customers", before, customers)
before = len(partners)
partners = strip_text(partners)
partners = clean_id(partners, ["partner_id"])
partners["partner_name"] = partners["partner_name"].str.title()
partners["phone"] = partners["phone"].apply(clean_phone)
partners["vehicle_type"] = standard(partners["vehicle_type"], vehicle_names)
partners["join_date"] = partners["join_date"].apply(to_date)
partners = drop_missing(
    partners, ["partner_id", "partner_name", "phone", "vehicle_type", "join_date"]
)
partners = remove_duplicates(partners, "partner_id")
show_result("Delivery partners", before, partners)
before = len(orders)
orders = strip_text(orders)
orders = clean_id(
    orders, ["order_id", "customer_id", "restaurant_id", "partner_id"]
)
orders["order_datetime"] = orders["order_datetime"].apply(to_date)

amount_text = (
    orders["order_amount"]
    .astype(str)
    .str.replace(r"Rs\.?|INR|,", "", regex=True)
    .str.strip()
)
orders["order_amount"] = pd.to_numeric(amount_text, errors="coerce")

delivery_text = orders["delivery_minutes"].astype(str)
orders["delivery_minutes"] = pd.to_numeric(
    delivery_text.str.extract(r"(\d+)")[0], errors="coerce"
)

orders["payment_mode"] = standard(orders["payment_mode"], payment_names)
orders["order_status"] = standard(orders["order_status"], status_names)
orders["coupon_code"] = orders["coupon_code"].str.upper()
orders = drop_missing(
    orders,
    [
        "order_id",
        "customer_id",
        "restaurant_id",
        "order_datetime",
        "order_amount",
        "payment_mode",
        "order_status",
    ],
)
orders = remove_duplicates(orders, "order_id")
show_result("Orders", before, orders)
before = len(feedback)
feedback = strip_text(feedback)
feedback = clean_id(feedback, ["feedback_id", "order_id"])
feedback["rating"] = pd.to_numeric(feedback["rating"], errors="coerce")
feedback["feedback_date"] = feedback["feedback_date"].apply(to_date)
feedback = drop_missing(
    feedback, ["feedback_id", "order_id", "rating", "feedback_date"]
)
feedback = feedback[feedback["rating"].between(1, 5)]
feedback = remove_duplicates(feedback, "feedback_id")
show_result("Feedback", before, feedback)
cleaned_tables = {
    "areas": areas,
    "restaurants": restaurants,
    "customers": customers,
    "delivery_partners": partners,
    "orders": orders,
    "feedback": feedback,
}

cleaning_file = output_folder / "Quick_Bite_Data_Cleaning.xlsx"

with pd.ExcelWriter(cleaning_file, engine="openpyxl") as writer:
    for sheet_name, table in cleaned_tables.items():
        table.to_excel(writer, sheet_name=sheet_name, index=False)

print("Cleaned tables saved to:", cleaning_file)

combined = orders.merge(
    customers[["customer_id", "customer_name"]],
    on="customer_id",
    how="left",
    validate="many_to_one"
)

combined = combined.merge(
    restaurants[["restaurant_id", "restaurant_name", "cuisine", "area_id"]],
    on="restaurant_id",
    how="left",
    validate="many_to_one"
)

combined = combined.merge(
    areas[["area_id", "area_name", "city"]],
    on="area_id",
    how="left",
    validate="many_to_one"
)

combined = combined.merge(
    partners[["partner_id", "partner_name", "vehicle_type"]],
    on="partner_id",
    how="left",
    validate="many_to_one"
)

feedback_for_join = (
    feedback.sort_values("feedback_date")
    .drop_duplicates(subset="order_id", keep="first")
)

combined = combined.merge(
    feedback_for_join[["order_id", "rating", "comment"]],
    on="order_id",
    how="left",
    validate="one_to_one"
)
combined["order_date"] = combined["order_datetime"].dt.date
combined["order_time"] = combined["order_datetime"].dt.strftime("%H:%M")
combined["order_day"] = combined["order_datetime"].dt.day_name()
combined["is_late"] = np.where(
    combined["delivery_minutes"] > 45,
    "Yes",
    "No"
)
combined.loc[
    combined["delivery_minutes"].isna(),
    "is_late"
] = "Not Available"
combined["partner_name"] = combined["partner_name"].fillna("Not Assigned")
combined["vehicle_type"] = combined["vehicle_type"].fillna("Not Assigned")
combined["coupon_code"] = combined["coupon_code"].fillna("No Coupon")
combined["comment"] = combined["comment"].fillna("No Comment Provided")
required_joined_columns = [
    "customer_name",
    "restaurant_name",
    "cuisine",
    "area_name",
    "city",
]

missing_link_counts = combined[required_joined_columns].isna().sum()
if missing_link_counts.any():
    print("Warning: Missing linked records detected:")
    print(missing_link_counts[missing_link_counts > 0])
    raise ValueError(
        "Some orders could not be matched to their customer, restaurant "
        "or area records. Check the corresponding IDs in the cleaned tables."
    )
final_columns = [
    "order_id",
    "order_date",
    "order_time",
    "order_day",
    "customer_name",
    "restaurant_name",
    "cuisine",
    "area_name",
    "city",
    "partner_name",
    "vehicle_type",
    "order_amount",
    "payment_mode",
    "coupon_code",
    "order_status",
    "delivery_minutes",
    "is_late",
    "rating",
    "comment",
]

combined = combined[final_columns]
if combined["order_id"].duplicated().any():
    raise ValueError("Duplicate order IDs found in the combined dataset.")

if len(combined) != len(orders):
    raise ValueError(
        f"Row count mismatch: {len(orders)} cleaned orders but "
        f"{len(combined)} combined rows."
    )
combined_file = output_folder / "Quick_Bite_Combined_Data.xlsx"

combined.to_excel(
    combined_file,
    sheet_name="combined_data",
    index=False
)

print("Combined rows (one per order):", len(combined))
print("Total columns:", len(combined.columns))
print("Missing delivery times:", combined["delivery_minutes"].isna().sum())
print("Missing ratings:", combined["rating"].isna().sum())
print("Missing comments:", combined["comment"].isna().sum())
print("Duplicate order IDs:", combined["order_id"].duplicated().sum())
print("Combined data saved to:", combined_file)

