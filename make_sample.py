"""Generate a synthetic Vietnamese retail sales CSV.

Each row is one product line (a product sold in a store on a date).
Revenue is quantity times the catalog unit price. The random seed is fixed,
so running this script again rewrites the same file.

Example
-------
    python make_sample.py
"""

from __future__ import annotations

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

# Week 1 starts on this Monday. Week 8 is 12 Oct 2026 through 18 Oct 2026.
WEEK1_START = date(2026, 8, 24)

# How busy each week is, and a quantity bump so the trend is not a flat line.
# Week 7 is quieter and week 8 is busier, which gives the report a clear
# week-over-week change. Used only when generating the default 8 weeks.
WEEK_WEIGHTS = [0.90, 0.95, 1.00, 1.05, 1.12, 1.08, 0.92, 1.25]
WEEK_QTY_FACTORS = [1.00, 1.00, 1.05, 1.00, 1.08, 1.05, 0.88, 1.18]

# region, store, salespeople who work at that store
STORES = [
    ("Miền Bắc", "VinMart Hoàn Kiếm", ["Nguyễn Văn An", "Trần Thị Bình"]),
    ("Miền Bắc", "WinMart Cầu Giấy", ["Lê Hoàng Cường", "Phạm Minh Đức"]),
    ("Miền Bắc", "Co.opmart Long Biên", ["Hoàng Thị Mai", "Võ Quốc Huy"]),
    ("Miền Trung", "Co.opmart Đà Nẵng", ["Đặng Thu Hà", "Bùi Thanh Tùng"]),
    ("Miền Trung", "VinMart Huế", ["Ngô Thị Lan", "Phan Văn Khoa"]),
    ("Miền Trung", "WinMart Nha Trang", ["Đỗ Minh Quân", "Lý Thị Ngọc"]),
    ("Miền Nam", "Co.opmart Nguyễn Đình Chiểu", ["Trịnh Văn Sơn", "Mai Thu Trang"]),
    ("Miền Nam", "VinMart Quận 7", ["Cao Hải Đăng", "Dương Thị Hoa"]),
    ("Miền Nam", "Big C Phú Mỹ Hưng", ["Huỳnh Văn Nam", "Tạ Thị Yến"]),
    ("Miền Nam", "WinMart Cần Thơ", ["Lương Minh Khang", "Vũ Thu Phương"]),
]

# weight: how often the product is chosen. max_qty: largest quantity on one line.
# Add a product by appending one dictionary.
PRODUCTS = [
    {"category": "Đồ uống", "product": "Trà xanh không độ", "unit_price": 10_000, "max_qty": 20, "weight": 4},
    {"category": "Đồ uống", "product": "Cà phê sữa đá lon", "unit_price": 15_000, "max_qty": 16, "weight": 3},
    {"category": "Đồ uống", "product": "Nước suối 500ml", "unit_price": 5_000, "max_qty": 30, "weight": 5},
    {"category": "Đồ uống", "product": "Nước ngọt có ga", "unit_price": 12_000, "max_qty": 18, "weight": 3},
    {"category": "Sữa & đồ lạnh", "product": "Sữa tươi tiệt trùng", "unit_price": 32_000, "max_qty": 12, "weight": 3},
    {"category": "Sữa & đồ lạnh", "product": "Sữa chua uống", "unit_price": 8_000, "max_qty": 20, "weight": 3},
    {"category": "Sữa & đồ lạnh", "product": "Phô mai lát", "unit_price": 45_000, "max_qty": 8, "weight": 2},
    {"category": "Thực phẩm khô", "product": "Mì gói", "unit_price": 4_000, "max_qty": 40, "weight": 5},
    {"category": "Thực phẩm khô", "product": "Gạo ST25 (1kg)", "unit_price": 28_000, "max_qty": 10, "weight": 2},
    {"category": "Thực phẩm khô", "product": "Nước mắm", "unit_price": 55_000, "max_qty": 6, "weight": 2},
    {"category": "Bánh kẹo", "product": "Bánh quy bơ", "unit_price": 25_000, "max_qty": 10, "weight": 2},
    {"category": "Bánh kẹo", "product": "Kẹo dẻo", "unit_price": 18_000, "max_qty": 12, "weight": 2},
    {"category": "Bánh kẹo", "product": "Snack khoai tây", "unit_price": 12_000, "max_qty": 16, "weight": 3},
    {"category": "Chăm sóc cá nhân", "product": "Dầu gội", "unit_price": 89_000, "max_qty": 4, "weight": 2},
    {"category": "Chăm sóc cá nhân", "product": "Sữa tắm", "unit_price": 75_000, "max_qty": 4, "weight": 2},
    {"category": "Chăm sóc cá nhân", "product": "Kem đánh răng", "unit_price": 28_000, "max_qty": 8, "weight": 2},
    {"category": "Gia dụng", "product": "Nước rửa chén", "unit_price": 35_000, "max_qty": 8, "weight": 2},
    {"category": "Gia dụng", "product": "Giấy vệ sinh", "unit_price": 42_000, "max_qty": 8, "weight": 2},
    {"category": "Gia dụng", "product": "Bột giặt", "unit_price": 95_000, "max_qty": 3, "weight": 1},
]

# Column names that report.py expects, in the order written to the CSV.
COLUMNS = [
    "date",
    "week",
    "region",
    "store",
    "product_category",
    "product",
    "quantity",
    "unit_price",
    "revenue",
    "salesperson",
]


def _week_profile(n_weeks: int) -> tuple[list[float], list[float]]:
    """Return (row weights, quantity factors) with one entry per week."""
    if n_weeks == len(WEEK_WEIGHTS):
        return WEEK_WEIGHTS, WEEK_QTY_FACTORS
    return [1.0] * n_weeks, [1.0] * n_weeks


def generate_sales(n_rows: int = 1000, n_weeks: int = 8, seed: int = 42) -> pd.DataFrame:
    """Build a reproducible sales table.

    Parameters
    ----------
    n_rows:
        How many product lines to create.
    n_weeks:
        Weeks are numbered 1 through n_weeks.
    seed:
        Random seed. The same seed always produces the same rows.
    """
    if n_rows < 1:
        raise ValueError("n_rows must be at least 1.")
    if n_weeks < 1:
        raise ValueError("n_weeks must be at least 1.")

    rng = random.Random(seed)
    weights, qty_factors = _week_profile(n_weeks)
    product_weights = [item["weight"] for item in PRODUCTS]
    week_numbers = list(range(1, n_weeks + 1))
    rows = []

    for _ in range(n_rows):
        week = rng.choices(week_numbers, weights=weights, k=1)[0]
        sold_on = WEEK1_START + timedelta(days=(week - 1) * 7 + rng.randint(0, 6))
        region, store, people = rng.choice(STORES)
        product = rng.choices(PRODUCTS, weights=product_weights, k=1)[0]
        base_qty = rng.randint(1, product["max_qty"])
        quantity = max(1, int(round(base_qty * qty_factors[week - 1])))
        unit_price = product["unit_price"]
        rows.append(
            {
                "date": sold_on.isoformat(),
                "week": week,
                "region": region,
                "store": store,
                "product_category": product["category"],
                "product": product["product"],
                "quantity": quantity,
                "unit_price": unit_price,
                "revenue": quantity * unit_price,
                "salesperson": rng.choice(people),
            }
        )

    sales = pd.DataFrame(rows, columns=COLUMNS)
    sales = sales.sort_values(["date", "region", "store", "product"]).reset_index(drop=True)
    return sales


def write_sales(sales: pd.DataFrame, output_path: str | Path) -> Path:
    """Write the sales table to a UTF-8 CSV and return the path."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    sales.to_csv(path, index=False, encoding="utf-8-sig")
    return path


def main() -> None:
    """Create sample_input/sales_data.csv (or a path passed with --output)."""
    parser = argparse.ArgumentParser(description="Generate a sample retail sales CSV.")
    parser.add_argument("--rows", type=int, default=1000, help="Number of product lines. Default: 1000.")
    parser.add_argument("--weeks", type=int, default=8, help="Number of weeks. Default: 8.")
    parser.add_argument("--seed", type=int, default=42, help="Random seed. Default: 42.")
    parser.add_argument(
        "--output",
        default="sample_input/sales_data.csv",
        help="Output CSV path. Default: sample_input/sales_data.csv.",
    )
    args = parser.parse_args()

    sales = generate_sales(n_rows=args.rows, n_weeks=args.weeks, seed=args.seed)
    path = write_sales(sales, args.output)
    counts = sales.groupby("week").size()
    count_text = ", ".join(f"{int(week)}={int(count)}" for week, count in counts.items())
    print(f"Wrote {len(sales)} rows to {path}")
    print(f"Dates {sales['date'].min()} to {sales['date'].max()}")
    print(f"Rows per week: {count_text}")
    print(f"Total revenue: {int(sales['revenue'].sum()):,} VND")


if __name__ == "__main__":
    main()
