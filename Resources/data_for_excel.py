import pandas as pd
import random
from datetime import datetime, timedelta

random.seed(99)

rows = []

customers = [
    "Ananya Dash", "Rahul  Sharma", "sneha patel", "AMIT KUMAR",
    "Priya Singh", "Rohit Verma ", "Neha  Gupta", "arjun reddy"
]

regions = ["East", " west", "NORTH", "south ", "Central"]
categories = ["Furniture", "Technology", "Office Supplies", "technology ", "furniture"]
sub_categories = ["Chairs", "Tables", "Phones", "Binders", "Storage", "Accessories"]

cities = ["Mumbai", "Delhi ", " bengaluru", "Chennai", "Hyderabad", "Pune"]
states = ["Maharashtra", " Delhi", "Karnataka ", "Tamil Nadu", "Telangana"]

products = [
    "Ergonomic Office Chair with lumbar support",
    "Standing Desk Adjustable Height",
    "Wireless Mouse – USB Receiver",
    "Laptop Backpack Water Resistant",
    "Noise Cancelling Headphones with Mic",
    "USB-C Hub Multiport Adapter"
]

sales_reps = ["Ravi", " Neha", "SURESH", "ravi ", "Amit", "NEHA"]

for i in range(1500):
    order_id = random.choice([f"ORD-{random.randint(1000,1200)}", f"ORD-{i+2000}"])
    customer_id = random.choice([random.randint(100, 400), random.randint(100, 150)])

    order_date = random.choice([
        (datetime.today() - timedelta(days=random.randint(1, 900))).strftime("%d-%m-%Y"),
        (datetime.today() - timedelta(days=random.randint(1, 900))).strftime("%Y/%m/%d"),
        "Not Available",
        ""
    ])

    ship_date = random.choice([
        (datetime.today() - timedelta(days=random.randint(1, 850))).strftime("%d-%m-%Y"),
        "NA",
        None
    ])

    quantity = random.choice([1, 2, 3, "2 ", " three", "", None])
    sales = random.choice(["₹1,250", "1200 ", "3,400", 2500, None])
    discount = random.choice([0, 0.1, "0.2 ", "ten%", None])
    profit = random.choice(["500", "-200 ", "₹300", None])

    rows.append([
        order_id,
        order_date,
        ship_date,
        customer_id,
        random.choice(customers),
        random.choice(regions),
        random.choice(states),
        random.choice(cities),
        random.choice(categories),
        random.choice(sub_categories),
        random.choice(products),
        quantity,
        sales,
        discount,
        profit,
        random.choice(sales_reps)
    ])

columns = [
    "Order ID", "Order Date", "Ship Date", "Customer ID", "Customer Name",
    "Region", "State", "City", "Category", "Sub-Category",
    "Product Name", "Quantity", "Sales", "Discount", "Profit", "Sales Rep"
]

df = pd.DataFrame(rows, columns=columns)

# Add duplicate rows intentionally
df = pd.concat([df, df.sample(60)], ignore_index=True)

# Save to Excel
df.to_excel("superstore_dirty_data.xlsx", index=False)

print("✅ superstore_dirty_data.xlsx generated with messy data")
