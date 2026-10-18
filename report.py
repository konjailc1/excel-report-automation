"""Build a formatted weekly sales workbook from a CSV file.

Pandas calculates the KPIs and pivot tables. openpyxl writes the sheets,
charts, number formats, and conditional formatting.

Example
-------
    from report import build_report

    build_report(
        "sample_input/sales_data.csv",
        week=8,
        output_path="output/weekly_report_week8.xlsx",
    )
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Columns the CSV must contain. Display labels are what Excel shows.
COLUMN_LABELS = {
    "date": "Date",
    "week": "Week",
    "region": "Region",
    "store": "Store",
    "product_category": "Product Category",
    "product": "Product",
    "quantity": "Quantity",
    "unit_price": "Unit Price (VND)",
    "revenue": "Revenue (VND)",
    "salesperson": "Salesperson",
}

# Geographic order for the three sample regions. Any other region is added after these.
REGION_ORDER = ["Miền Bắc", "Miền Trung", "Miền Nam"]

VND_FORMAT = "#,##0"
PERCENT_FORMAT = "+0.0%;-0.0%;0.0%"
SHARE_FORMAT = "0.0%"
DATE_FORMAT = "YYYY-MM-DD"

NAVY = "1F4E79"
NAVY_LIGHT = "2E75B6"
WHITE = "FFFFFF"
INK = "1F1F1F"
MUTED = "595959"
GRID = "D9D9D9"
TOTAL_FILL = "D6DCE4"
ZEBRA = "F7F9FB"
GOOD_FILL_COLOR = "C6EFCE"
GOOD_FONT_COLOR = "006100"
BAD_FILL_COLOR = "FFC7CE"
BAD_FONT_COLOR = "9C0006"

CARD_FILLS = ["D6EAF8", "E8F6EF", "FDEBD0", "F5EEF8"]
COMMENT_FILL_COLOR = "F4F8FB"

# Words that change color in the commentary. The figures next to them are shared.
CHANGE_WORDS = {
    "en": {"up": "increased", "down": "decreased", "flat": "unchanged"},
    "vi": {"up": "tăng", "down": "giảm", "flat": "không đổi"},
}


def _font(size: int = 11, bold: bool = False, color: str = INK) -> Font:
    """Calibri font used everywhere so the workbook looks consistent."""
    return Font(name="Calibri", size=size, bold=bold, color=color)


def _fill(color: str) -> PatternFill:
    """Solid cell fill from a hex color such as '1F4E79'."""
    return PatternFill(fill_type="solid", fgColor=color)


def _border(color: str = GRID, style: str = "thin") -> Border:
    """Box border around one cell."""
    side = Side(style=style, color=color)
    return Border(left=side, right=side, top=side, bottom=side)


THIN = _border()
CARD_BORDER = _border(NAVY, "medium")
HEADER_FONT = _font(11, bold=True, color=WHITE)
HEADER_FILL = _fill(NAVY)
TITLE_FONT = _font(20, bold=True, color=WHITE)
SUBTITLE_FONT = _font(12, color=WHITE)
LABEL_FONT = _font(10, bold=True, color=NAVY)
VALUE_FONT = _font(20, bold=True, color=NAVY)
HINT_FONT = _font(9, color=MUTED)
SECTION_FONT = _font(14, bold=True, color=NAVY)
BODY_FONT = _font(11)
TOTAL_FONT = _font(11, bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", indent=1)
RIGHT = Alignment(horizontal="right", vertical="center")

# Excel uses bgColor for a solid conditional-format fill, so both colors match.
GOOD_FILL = PatternFill(fill_type="solid", fgColor=GOOD_FILL_COLOR, bgColor=GOOD_FILL_COLOR)
BAD_FILL = PatternFill(fill_type="solid", fgColor=BAD_FILL_COLOR, bgColor=BAD_FILL_COLOR)
GOOD_FONT = Font(color=GOOD_FONT_COLOR, bold=True)
BAD_FONT = Font(color=BAD_FONT_COLOR, bold=True)


def load_sales(csv_path: str | Path) -> pd.DataFrame:
    """Read the sales CSV and check that the required columns exist.

    Dates are parsed, and numeric columns are converted to numbers.
    The revenue column is used as stored (it is not recalculated), so a
    later discount column can change revenue without breaking the loader.
    """
    path = Path(csv_path)
    if not path.is_file():
        raise FileNotFoundError(f"Sales file not found: {path}")

    try:
        sales = pd.read_csv(path, encoding="utf-8-sig")
    except pd.errors.EmptyDataError as exc:
        raise ValueError(f"The sales file is empty: {path}") from exc

    sales.columns = [str(name).strip() for name in sales.columns]
    missing = [name for name in COLUMN_LABELS if name not in sales.columns]
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))

    sales = sales.loc[:, list(COLUMN_LABELS)].copy()
    if sales.empty:
        raise ValueError(f"The sales file has no rows: {path}")

    sales["date"] = pd.to_datetime(sales["date"], errors="coerce")
    if sales["date"].isna().any():
        raise ValueError("The date column has values that are not dates. Use YYYY-MM-DD.")

    for column in ("week", "quantity", "unit_price", "revenue"):
        sales[column] = pd.to_numeric(sales[column], errors="coerce")
        if sales[column].isna().any():
            raise ValueError(f"The {column} column must contain numbers only.")

    sales["week"] = sales["week"].astype(int)
    sales["quantity"] = sales["quantity"].astype(int)
    sales["unit_price"] = sales["unit_price"].astype(int)
    sales["revenue"] = sales["revenue"].astype(int)
    return sales


def rows_through_week(sales: pd.DataFrame, week: int) -> pd.DataFrame:
    """Keep week 1 through `week` so a report never includes later weeks."""
    return sales.loc[sales["week"] <= week].copy()


def percent_change(current: float, previous: float | None) -> float | None:
    """Return (current - previous) / previous.

    Returns None when there is no previous number or the previous number is 0,
    because a percent change would not mean anything in those cases.
    """
    if previous is None or previous == 0:
        return None
    return (current - previous) / previous


def compute_kpis(sales: pd.DataFrame, week: int) -> dict:
    """Summarize one week and compare it with the previous week.

    An "order" here is one row: one product line. Average order value is
    revenue divided by the number of rows, rounded to the nearest dong.

    The previous week must already be in `sales`. If it is missing, the
    week-over-week figures are None.
    """
    current = sales.loc[sales["week"] == week]
    if current.empty:
        raise ValueError(f"No sales rows for week {week}.")

    previous = sales.loc[sales["week"] == week - 1]
    revenue = int(current["revenue"].sum())
    orders = int(len(current))
    avg_order_value = int(round(revenue / orders))

    previous_revenue = None
    previous_orders = None
    previous_avg = None
    if not previous.empty:
        previous_revenue = int(previous["revenue"].sum())
        previous_orders = int(len(previous))
        previous_avg = int(round(previous_revenue / previous_orders))

    return {
        "week": int(week),
        "start_date": current["date"].min(),
        "end_date": current["date"].max(),
        "revenue": revenue,
        "orders": orders,
        "avg_order_value": avg_order_value,
        "wow_change": percent_change(revenue, previous_revenue),
        "previous_revenue": previous_revenue,
        "previous_orders": previous_orders,
        "previous_avg_order_value": previous_avg,
        "orders_change": percent_change(orders, previous_orders),
        "avg_order_value_change": percent_change(avg_order_value, previous_avg),
    }


def sort_regions(names) -> list[str]:
    """Order regions North, Central, South, then any other name A to Z."""
    present = {str(name) for name in names}
    ordered = [name for name in REGION_ORDER if name in present]
    extras = sorted(name for name in present if name not in REGION_ORDER)
    return ordered + extras


def revenue_by_region(sales: pd.DataFrame, week: int) -> pd.DataFrame:
    """Revenue and share of revenue for each region in one week.

    Regions with no sales that week are included as 0 if they appear somewhere
    else in `sales`, so the summary does not drop a region just because it was quiet.
    """
    current = sales.loc[sales["week"] == week]
    totals = current.groupby("region")["revenue"].sum()
    grand = int(totals.sum()) if len(totals) else 0
    rows = []
    for region in sort_regions(sales["region"].unique()):
        revenue = int(totals[region]) if region in totals.index else 0
        share = (revenue / grand) if grand else 0
        rows.append({"region": region, "revenue": revenue, "share": share})
    return pd.DataFrame(rows)


def pivot_revenue(
    sales: pd.DataFrame,
    index_column: str,
    index_label: str,
    index_order: list[str],
) -> pd.DataFrame:
    """Sum revenue with one column per week, plus a total row and total column.

    This uses pandas.pivot_table (margins=True), which is the same idea as an
    Excel pivot: each intersection is the sum of revenue, and Total is the sum
    of the row or column.
    """
    labeled = sales.copy()
    labeled["week_label"] = labeled["week"].map(lambda week: f"Week {int(week)}")
    table = pd.pivot_table(
        labeled,
        index=index_column,
        columns="week_label",
        values="revenue",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Total",
    )
    if isinstance(table.columns, pd.MultiIndex):
        table.columns = [column[-1] for column in table.columns]

    week_columns = [column for column in table.columns if column != "Total"]
    week_columns = sorted(week_columns, key=lambda name: int(str(name).replace("Week ", "")))
    table = table.loc[:, week_columns + ["Total"]]

    body = [name for name in index_order if name in table.index]
    already = set(body)
    body.extend(name for name in table.index if name not in already and name != "Total")
    table = table.reindex(body + ["Total"])
    table.index.name = index_label
    return table


def weekly_totals(sales: pd.DataFrame) -> pd.DataFrame:
    """One row per week: revenue and the number of product lines (orders)."""
    totals = (
        sales.groupby("week", as_index=False)
        .agg(revenue=("revenue", "sum"), orders=("revenue", "size"))
        .sort_values("week")
    )
    totals["week"] = totals["week"].astype(int)
    totals["revenue"] = totals["revenue"].astype(int)
    totals["orders"] = totals["orders"].astype(int)
    return totals.reset_index(drop=True)


def top_products(sales: pd.DataFrame, week: int, limit: int = 10) -> pd.DataFrame:
    """Best-selling products in one week, highest revenue first.

    Rank 1 is the first row. The bar chart flips its axis so that row still
    appears at the top of the chart.
    """
    current = sales.loc[sales["week"] == week]
    grouped = (
        current.groupby("product", as_index=False)
        .agg(revenue=("revenue", "sum"), quantity=("quantity", "sum"))
        .sort_values(["revenue", "product"], ascending=[False, True])
        .head(limit)
    )
    grouped["revenue"] = grouped["revenue"].astype(int)
    grouped["quantity"] = grouped["quantity"].astype(int)
    grouped.insert(0, "rank", range(1, len(grouped) + 1))
    return grouped.reset_index(drop=True)


def _vnd_text(amount: int) -> str:
    """Format dong with thousands separators and no decimals (#,##0)."""
    return f"{int(amount):,} VND"


def _pct_text(value: float) -> str:
    """Format a ratio as a percent with one decimal, such as 66.5%."""
    return f"{abs(float(value)):.1%}"


def _direction(change: float | None) -> str | None:
    """Classify a percent change as up, down, flat, or missing."""
    if change is None:
        return None
    if change > 0:
        return "up"
    if change < 0:
        return "down"
    return "flat"


def _commentary_facts(df: pd.DataFrame, week: int) -> dict:
    """Calculate every figure used in the commentary.

    Build this once per language. The strings here are the only place numbers
    are formatted, so Vietnamese and English cannot disagree.
    """
    week = int(week)
    kpis = compute_kpis(df, week)
    regions = revenue_by_region(df, week)
    current = df.loc[df["week"] == week]

    ranked = regions.sort_values(["revenue", "region"], ascending=[False, True])
    best = ranked.iloc[0]
    weak = ranked.iloc[-1]
    revenue = int(kpis["revenue"])

    categories = (
        current.groupby("product_category", as_index=False)["revenue"]
        .sum()
        .sort_values(["revenue", "product_category"], ascending=[False, True])
    )
    top_category = categories.iloc[0]
    category_revenue = int(top_category["revenue"])
    category_share = (category_revenue / revenue) if revenue else 0

    product = top_products(df, week, limit=1).iloc[0]
    product_revenue = int(product["revenue"])
    best_share = float(best["share"])

    return {
        "week": week,
        "previous_week": None if kpis["previous_revenue"] is None else week - 1,
        "revenue_text": _vnd_text(revenue),
        "wow_direction": _direction(kpis["wow_change"]),
        "wow_pct": None if kpis["wow_change"] is None else _pct_text(kpis["wow_change"]),
        "best_region": str(best["region"]),
        "best_revenue_text": _vnd_text(best["revenue"]),
        "best_share_text": _pct_text(best_share),
        "best_share": best_share,
        "weak_region": str(weak["region"]),
        "weak_revenue_text": _vnd_text(weak["revenue"]),
        "weak_share_text": _pct_text(weak["share"]),
        "region_count": int(len(regions)),
        "regions_level": int(best["revenue"]) == int(weak["revenue"]),
        "top_category": str(top_category["product_category"]),
        "top_category_revenue_text": _vnd_text(category_revenue),
        "top_category_share_text": _pct_text(category_share),
        "top_product": str(product["product"]),
        "top_product_revenue_text": _vnd_text(product_revenue),
        "aov_text": _vnd_text(kpis["avg_order_value"]),
        "previous_aov_text": None
        if kpis["previous_avg_order_value"] is None
        else _vnd_text(kpis["previous_avg_order_value"]),
        "aov_direction": _direction(kpis["avg_order_value_change"]),
        "aov_pct": None
        if kpis["avg_order_value_change"] is None
        else _pct_text(kpis["avg_order_value_change"]),
    }


def _change_span(direction: str, percent: str | None, lang: str) -> tuple[str, str]:
    """Return the increase/decrease words and the tone used to color them."""
    word = CHANGE_WORDS[lang][direction]
    if direction == "flat" or not percent:
        return word, "text"
    return f"{word} {percent}", direction


def _recommendation_kind(facts: dict) -> str:
    """Pick one recommendation from the same figures both languages use."""
    level = facts["regions_level"] or facts["region_count"] == 1
    if facts["wow_direction"] is None:
        return "baseline_single" if level else "baseline"
    if facts["wow_direction"] == "down":
        return "recover_single" if level else "recover"
    if facts["aov_direction"] == "down":
        return "basket"
    if not level and facts["best_share"] >= 0.45:
        return "rebalance"
    return "restock"


def _bullet(*parts: tuple[str, str]) -> list[tuple[str, str]]:
    """One commentary line. Each part is (text, tone). Tone is text, up, or down."""
    return [("• ", "text"), *parts]


def _commentary_en(facts: dict) -> list[list[tuple[str, str]]]:
    """Five English bullets. Numbers are read from `facts`, not typed in."""
    week = facts["week"]
    previous = facts["previous_week"]
    bullets = []

    if facts["wow_direction"] is None:
        bullets.append(_bullet((
            f"Week {week} revenue was {facts['revenue_text']}. "
            "There is no earlier week to compare.",
            "text",
        )))
    elif facts["wow_direction"] == "flat":
        bullets.append(_bullet((
            f"Week {week} revenue was {facts['revenue_text']}. "
            f"It was unchanged from week {previous}.",
            "text",
        )))
    else:
        change, tone = _change_span(facts["wow_direction"], facts["wow_pct"], "en")
        bullets.append(_bullet(
            (f"Week {week} revenue was {facts['revenue_text']}. It ", "text"),
            (change, tone),
            (f" from week {previous}.", "text"),
        ))

    if facts["region_count"] == 1:
        bullets.append(_bullet((
            f"{facts['best_region']} was the only region with sales, at "
            f"{facts['best_revenue_text']} ({facts['best_share_text']} of revenue).",
            "text",
        )))
    elif facts["regions_level"]:
        bullets.append(_bullet((
            "Revenue was level across regions. The high and the low were both "
            f"{facts['best_revenue_text']} ({facts['best_share_text']} of revenue).",
            "text",
        )))
    else:
        bullets.append(_bullet((
            f"{facts['best_region']} led with {facts['best_revenue_text']} "
            f"({facts['best_share_text']} of revenue). "
            f"{facts['weak_region']} was weakest at {facts['weak_revenue_text']} "
            f"({facts['weak_share_text']} of revenue).",
            "text",
        )))

    bullets.append(_bullet((
        f"Top category was {facts['top_category']} at {facts['top_category_revenue_text']} "
        f"({facts['top_category_share_text']} of revenue). "
        f"Top product was {facts['top_product']} at {facts['top_product_revenue_text']}.",
        "text",
    )))

    if facts["aov_direction"] is None:
        bullets.append(_bullet((
            f"Average order value was {facts['aov_text']}. "
            "This is the first week, so there is no trend yet.",
            "text",
        )))
    elif facts["aov_direction"] == "flat":
        bullets.append(_bullet((
            f"Average order value was {facts['aov_text']}. "
            f"It was unchanged from {facts['previous_aov_text']} in week {previous}.",
            "text",
        )))
    else:
        change, tone = _change_span(facts["aov_direction"], facts["aov_pct"], "en")
        bullets.append(_bullet(
            (f"Average order value was {facts['aov_text']}. It ", "text"),
            (change, tone),
            (f" from {facts['previous_aov_text']} in week {previous}.", "text"),
        ))

    kind = _recommendation_kind(facts)
    if kind == "baseline":
        text = (
            f"Treat week {week} as the baseline. Next week, compare "
            f"{facts['weak_region']} ({facts['weak_share_text']}) with "
            f"{facts['best_region']} ({facts['best_share_text']})."
        )
        bullets.append(_bullet((text, "text")))
    elif kind == "baseline_single":
        text = (
            f"Treat week {week} as the baseline. Next week, watch whether "
            f"{facts['top_product']} stays the best seller."
        )
        bullets.append(_bullet((text, "text")))
    elif kind == "recover":
        change, tone = _change_span("down", facts["wow_pct"], "en")
        bullets.append(_bullet(
            ("Revenue ", "text"),
            (change, tone),
            (
                f". Check stock and staffing in {facts['weak_region']} "
                f"({facts['weak_share_text']}) and feature {facts['top_product']} there.",
                "text",
            ),
        ))
    elif kind == "recover_single":
        change, tone = _change_span("down", facts["wow_pct"], "en")
        bullets.append(_bullet(
            ("Revenue ", "text"),
            (change, tone),
            (
                f". Check stock of {facts['top_category']}, and give "
                f"{facts['top_product']} a more visible spot.",
                "text",
            ),
        ))
    elif kind == "basket":
        change, tone = _change_span("down", facts["aov_pct"], "en")
        bullets.append(_bullet(
            ("Average order value ", "text"),
            (change, tone),
            (
                f". Pair {facts['top_product']} with another {facts['top_category']} "
                "item to bring the basket back up.",
                "text",
            ),
        ))
    elif kind == "rebalance":
        bullets.append(_bullet((
            f"{facts['best_region']} already accounts for {facts['best_share_text']} of revenue. "
            f"Put the next promotion in {facts['weak_region']} ({facts['weak_share_text']}) "
            "so sales are less concentrated.",
            "text",
        )))
    else:
        bullets.append(_bullet((
            f"Restock {facts['top_product']} in {facts['best_region']}. "
            "That product is the best seller, and that region leads the week.",
            "text",
        )))
    return bullets


def _commentary_vi(facts: dict) -> list[list[tuple[str, str]]]:
    """Five Vietnamese bullets. Numbers are the same strings as the English ones."""
    week = facts["week"]
    previous = facts["previous_week"]
    bullets = []

    if facts["wow_direction"] is None:
        bullets.append(_bullet((
            f"Doanh thu tuần {week} đạt {facts['revenue_text']}. "
            "Chưa có tuần trước để so sánh.",
            "text",
        )))
    else:
        change, tone = _change_span(facts["wow_direction"], facts["wow_pct"], "vi")
        bullets.append(_bullet(
            (f"Doanh thu tuần {week} đạt {facts['revenue_text']}, ", "text"),
            (change, tone),
            (f" so với tuần {previous}.", "text"),
        ))

    if facts["region_count"] == 1:
        bullets.append(_bullet((
            f"{facts['best_region']} là miền duy nhất có doanh thu, đạt "
            f"{facts['best_revenue_text']} ({facts['best_share_text']} doanh thu).",
            "text",
        )))
    elif facts["regions_level"]:
        bullets.append(_bullet((
            "Doanh thu các miền ngang nhau. Mức cao và mức thấp đều là "
            f"{facts['best_revenue_text']} ({facts['best_share_text']} doanh thu).",
            "text",
        )))
    else:
        bullets.append(_bullet((
            f"{facts['best_region']} dẫn đầu với {facts['best_revenue_text']} "
            f"({facts['best_share_text']} doanh thu). "
            f"{facts['weak_region']} thấp nhất với {facts['weak_revenue_text']} "
            f"({facts['weak_share_text']} doanh thu).",
            "text",
        )))

    bullets.append(_bullet((
        f"Nhóm hàng dẫn đầu là {facts['top_category']} với {facts['top_category_revenue_text']} "
        f"({facts['top_category_share_text']} doanh thu). "
        f"Sản phẩm bán chạy nhất là {facts['top_product']} với {facts['top_product_revenue_text']}.",
        "text",
    )))

    if facts["aov_direction"] is None:
        bullets.append(_bullet((
            f"Giá trị đơn trung bình đạt {facts['aov_text']}. "
            "Đây là tuần đầu nên chưa có xu hướng.",
            "text",
        )))
    else:
        change, tone = _change_span(facts["aov_direction"], facts["aov_pct"], "vi")
        bullets.append(_bullet(
            (f"Giá trị đơn trung bình đạt {facts['aov_text']}, ", "text"),
            (change, tone),
            (f" so với {facts['previous_aov_text']} ở tuần {previous}.", "text"),
        ))

    kind = _recommendation_kind(facts)
    if kind == "baseline":
        text = (
            f"Lấy tuần {week} làm mốc so sánh. Tuần sau hãy đối chiếu "
            f"{facts['weak_region']} ({facts['weak_share_text']}) với "
            f"{facts['best_region']} ({facts['best_share_text']})."
        )
        bullets.append(_bullet((text, "text")))
    elif kind == "baseline_single":
        text = (
            f"Lấy tuần {week} làm mốc so sánh. Tuần sau hãy xem "
            f"{facts['top_product']} còn là sản phẩm bán chạy nhất hay không."
        )
        bullets.append(_bullet((text, "text")))
    elif kind == "recover":
        change, tone = _change_span("down", facts["wow_pct"], "vi")
        bullets.append(_bullet(
            ("Doanh thu ", "text"),
            (change, tone),
            (
                f". Nên kiểm tra hàng và nhân sự ở {facts['weak_region']} "
                f"({facts['weak_share_text']}) và đặt {facts['top_product']} ở vị trí dễ thấy.",
                "text",
            ),
        ))
    elif kind == "recover_single":
        change, tone = _change_span("down", facts["wow_pct"], "vi")
        bullets.append(_bullet(
            ("Doanh thu ", "text"),
            (change, tone),
            (
                f". Nên kiểm tra tồn kho nhóm {facts['top_category']} và đặt "
                f"{facts['top_product']} ở vị trí dễ thấy hơn.",
                "text",
            ),
        ))
    elif kind == "basket":
        change, tone = _change_span("down", facts["aov_pct"], "vi")
        bullets.append(_bullet(
            ("Giá trị đơn ", "text"),
            (change, tone),
            (
                f". Nên bán kèm {facts['top_product']} với một món khác trong nhóm "
                f"{facts['top_category']} để kéo giá trị đơn lên.",
                "text",
            ),
        ))
    elif kind == "rebalance":
        bullets.append(_bullet((
            f"{facts['best_region']} đã chiếm {facts['best_share_text']} doanh thu. "
            f"Nên đưa chương trình kế tiếp sang {facts['weak_region']} "
            f"({facts['weak_share_text']}) để doanh số bớt dồn một miền.",
            "text",
        )))
    else:
        bullets.append(_bullet((
            f"Nên ưu tiên nhập {facts['top_product']} cho {facts['best_region']}. "
            "Đây là sản phẩm bán chạy nhất, và miền này đang dẫn đầu tuần.",
            "text",
        )))
    return bullets


def build_commentary(df: pd.DataFrame, week: int, lang: str) -> list[list[tuple[str, str]]]:
    """Build the summary bullets in Vietnamese or English.

    Parameters
    ----------
    df:
        Sales rows. Include the previous week if the text should show a change.
    week:
        The week being reported.
    lang:
        "vi" or "en".

    Returns
    -------
    Five bullets. Each bullet is a list of (text, tone) pieces. Tone is
    "text", "up", or "down". "up" is drawn green and "down" is drawn red.
    Both languages call `_commentary_facts`, so they show the same numbers.
    """
    if lang not in ("vi", "en"):
        raise ValueError('lang must be "vi" or "en".')
    facts = _commentary_facts(df, week)
    if lang == "vi":
        return _commentary_vi(facts)
    return _commentary_en(facts)


def _as_int(value) -> int:
    """Round a pivot or KPI number to a whole dong."""
    return int(round(float(value)))


def _paint(ws, row: int, column: int, value, font, fill, alignment, border, number_format=None):
    """Write one cell and apply the usual formatting in one place.

    Pass fill=None to leave the background empty. Raw Data does this so the
    Excel Table row stripes are not covered by a cell color.
    """
    cell = ws.cell(row, column, value)
    cell.font = font
    if fill is not None:
        cell.fill = fill
    cell.alignment = alignment
    cell.border = border
    if number_format:
        cell.number_format = number_format
    return cell


def _merge(ws, start_row, start_col, end_row, end_col, value, font, fill, alignment, border=None):
    """Merge a rectangle, fill every cell, and put `value` in the top-left cell.

    Excel only shows the top-left value of a merge. Filling the other cells
    avoids white gaps inside the merged area.
    """
    for row in range(start_row, end_row + 1):
        for column in range(start_col, end_col + 1):
            cell = ws.cell(row, column, None)
            cell.fill = fill
            cell.font = font
            cell.alignment = alignment
            if border is not None:
                cell.border = border
    ws.cell(start_row, start_col).value = value
    if start_row != end_row or start_col != end_col:
        ws.merge_cells(
            start_row=start_row,
            start_column=start_col,
            end_row=end_row,
            end_column=end_col,
        )
    return ws.cell(start_row, start_col)


def _style_header_row(ws, row: int, start_col: int, end_col: int) -> None:
    """Navy header with white centered text."""
    for column in range(start_col, end_col + 1):
        cell = ws.cell(row, column)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN
    ws.row_dimensions[row].height = 22


def add_wow_colors(ws, cell_range: str) -> None:
    """Green when a change is above 0, red when it is below 0.

    Zero stays uncolored. The rule requires a number, because Excel treats
    text (such as N/A or an em dash) as greater than zero.
    """
    anchor = cell_range.split(":")[0]
    ws.conditional_formatting.add(
        cell_range,
        FormulaRule(
            formula=[f"AND(ISNUMBER({anchor}),{anchor}>0)"],
            fill=GOOD_FILL,
            font=GOOD_FONT,
        ),
    )
    ws.conditional_formatting.add(
        cell_range,
        FormulaRule(
            formula=[f"AND(ISNUMBER({anchor}),{anchor}<0)"],
            fill=BAD_FILL,
            font=BAD_FONT,
        ),
    )


def autosize_columns(ws, ignore_rows: set[int] | None = None, min_width: int = 10, max_width: int = 42) -> None:
    """Set each column width from the longest value in that column.

    `ignore_rows` skips banner titles. A long title in column A would otherwise
    make that column as wide as the whole title.
    Number formats are estimated so a value like 1234567 is sized as '1,234,567'
    rather than as the shorter raw digits.
    """
    skipped = ignore_rows or set()
    for column_cells in ws.columns:
        letter = get_column_letter(column_cells[0].column)
        longest = 0
        for cell in column_cells:
            if cell.row in skipped or cell.value is None:
                continue
            text = str(cell.value)
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                if cell.number_format == VND_FORMAT:
                    text = f"{cell.value:,.0f}"
                elif "%" in str(cell.number_format):
                    text = f"{cell.value:.1%}"
            longest = max(longest, len(text))
        if longest == 0:
            continue
        ws.column_dimensions[letter].width = min(max(longest + 3, min_width), max_width)


def _prepare_page(ws, fit_width: bool = True) -> None:
    """Landscape A4. Most sheets shrink to one page wide."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1 if fit_width else 0
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = fit_width
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.page_setup.horizontalCentered = True


def _write_change(ws, row: int, column: int, change: float | None):
    """Write a percent change, or an em dash when it cannot be calculated."""
    if change is None:
        return _paint(ws, row, column, "—", BODY_FONT, _fill(WHITE), CENTER, THIN)
    return _paint(
        ws, row, column, change, BODY_FONT, _fill(WHITE), CENTER, THIN, PERCENT_FORMAT
    )


def _inline_font(tone: str) -> InlineFont:
    """Plain text, or the green/red used for an increase or a decrease."""
    if tone == "up":
        return InlineFont(rFont="Calibri", sz=11, b=True, color="FF006100")
    if tone == "down":
        return InlineFont(rFont="Calibri", sz=11, b=True, color="FF9C0006")
    return InlineFont(rFont="Calibri", sz=11, color="FF1F1F1F")


def _rich_text(segments: list[tuple[str, str]]) -> CellRichText:
    """Join commentary pieces into one Excel rich-text value."""
    blocks = []
    for text, tone in segments:
        if text:
            blocks.append(TextBlock(_inline_font(tone), text))
    return CellRichText(*blocks)


def _write_rich_bullet(ws, row: int, start_col: int, end_col: int, segments) -> None:
    """Write one wrapped bullet across a merged box."""
    fill = _fill(COMMENT_FILL_COLOR)
    alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for column in range(start_col, end_col + 1):
        cell = ws.cell(row, column, None)
        cell.fill = fill
        cell.border = THIN
        cell.alignment = alignment
    if start_col != end_col:
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = ws.cell(row, start_col)
    cell.value = _rich_text(segments)
    cell.alignment = alignment
    cell.fill = fill


def _bullet_lines(segments, width: int) -> int:
    """Estimate how many wrapped lines a bullet needs."""
    length = sum(len(text) for text, _tone in segments)
    if length <= 0:
        return 1
    return (length + width - 1) // width


def _write_commentary(ws, sales: pd.DataFrame, week: int, start_row: int) -> None:
    """Two side-by-side boxes: Vietnamese on the left, English on the right."""
    vietnamese = build_commentary(sales, week, "vi")
    english = build_commentary(sales, week, "en")
    if len(vietnamese) != len(english):
        raise ValueError("Vietnamese and English commentary must have the same bullets.")

    _merge(
        ws, start_row, 2, start_row, 5,
        "Nhận xét (Tiếng Việt)", HEADER_FONT, HEADER_FILL, LEFT, THIN,
    )
    _merge(
        ws, start_row, 6, start_row, 9,
        "Commentary (English)", HEADER_FONT, HEADER_FILL, LEFT, THIN,
    )
    ws.row_dimensions[start_row].height = 22

    for index, (vi_bullet, en_bullet) in enumerate(zip(vietnamese, english)):
        row = start_row + 1 + index
        _write_rich_bullet(ws, row, 2, 5, vi_bullet)
        _write_rich_bullet(ws, row, 6, 9, en_bullet)
        # Both boxes are about 70 characters wide. Size the row from the longer bullet
        # and leave a little padding so wrapped lines do not touch the border.
        lines = max(_bullet_lines(vi_bullet, 62), _bullet_lines(en_bullet, 62))
        ws.row_dimensions[row].height = max(36, lines * 18 + 6)


def write_summary(ws, sales: pd.DataFrame, kpis: dict, regions: pd.DataFrame) -> None:
    """KPI cards, comparison tables, and bilingual commentary for the week."""
    ws.title = "Summary"
    ws.sheet_properties.tabColor = NAVY
    _prepare_page(ws)

    widths = {"A": 3, "B": 24, "C": 18, "D": 18, "E": 16, "F": 20, "G": 24, "H": 18, "I": 14}
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width

    week = kpis["week"]
    start = pd.Timestamp(kpis["start_date"]).strftime("%d %b %Y")
    end = pd.Timestamp(kpis["end_date"]).strftime("%d %b %Y")
    _merge(ws, 1, 1, 1, 9, f"Weekly Sales Report — Week {week}", TITLE_FONT, _fill(NAVY), LEFT)
    _merge(
        ws,
        2,
        1,
        2,
        9,
        f"{start} to {end}    |    All figures in VND",
        SUBTITLE_FONT,
        _fill(NAVY_LIGHT),
        LEFT,
    )
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 10

    previous_hint = f"vs week {week - 1}" if kpis["wow_change"] is not None else "no previous week"
    cards = [
        ("TOTAL REVENUE", kpis["revenue"], VND_FORMAT, "VND"),
        ("ORDERS", kpis["orders"], VND_FORMAT, "product lines"),
        ("AVG ORDER VALUE", kpis["avg_order_value"], VND_FORMAT, "VND per product line"),
        ("WEEK OVER WEEK", kpis["wow_change"], PERCENT_FORMAT, previous_hint),
    ]
    # Each card is two columns wide: B-C, D-E, F-G, H-I.
    card_columns = [2, 4, 6, 8]
    wow_cell = None
    for index, (label, value, number_format, hint) in enumerate(cards):
        column = card_columns[index]
        fill = _fill(CARD_FILLS[index])
        for row in range(4, 7):
            for extra in (0, 1):
                cell = ws.cell(row, column + extra)
                cell.fill = fill
                cell.border = CARD_BORDER
        _merge(ws, 4, column, 4, column + 1, label, LABEL_FONT, fill, CENTER, CARD_BORDER)
        shown = "N/A" if value is None else value
        format_to_use = None if value is None else number_format
        value_cell = _merge(
            ws, 5, column, 5, column + 1, shown, VALUE_FONT, fill, CENTER, CARD_BORDER
        )
        if format_to_use:
            value_cell.number_format = format_to_use
        _merge(ws, 6, column, 6, column + 1, hint, HINT_FONT, fill, CENTER, CARD_BORDER)
        if index == 3:
            wow_cell = value_cell

    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 36
    ws.row_dimensions[6].height = 18
    if wow_cell is not None and kpis["wow_change"] is not None:
        add_wow_colors(ws, wow_cell.coordinate)

    _merge(ws, 8, 2, 8, 5, "Compared with the previous week", SECTION_FONT, _fill(WHITE), LEFT)
    previous_header = f"Week {week - 1}" if kpis["previous_revenue"] is not None else "Previous week"
    headers = ["Metric", f"Week {week}", previous_header, "Change"]
    for offset, header in enumerate(headers):
        _paint(ws, 9, 2 + offset, header, HEADER_FONT, HEADER_FILL, CENTER, THIN)
    ws.row_dimensions[9].height = 22

    comparison_rows = [
        ("Revenue", kpis["revenue"], kpis["previous_revenue"], kpis["wow_change"]),
        ("Orders", kpis["orders"], kpis["previous_orders"], kpis["orders_change"]),
        (
            "Avg order value",
            kpis["avg_order_value"],
            kpis["previous_avg_order_value"],
            kpis["avg_order_value_change"],
        ),
    ]
    for offset, (metric, current, previous, change) in enumerate(comparison_rows):
        row = 10 + offset
        zebra = _fill(ZEBRA) if offset % 2 else _fill(WHITE)
        _paint(ws, row, 2, metric, BODY_FONT, zebra, LEFT, THIN)
        _paint(ws, row, 3, current, BODY_FONT, zebra, RIGHT, THIN, VND_FORMAT)
        if previous is None:
            _paint(ws, row, 4, "—", BODY_FONT, zebra, CENTER, THIN)
        else:
            _paint(ws, row, 4, previous, BODY_FONT, zebra, RIGHT, THIN, VND_FORMAT)
        change_cell = _write_change(ws, row, 5, change)
        change_cell.fill = zebra
    add_wow_colors(ws, "E10:E12")

    _merge(ws, 8, 7, 8, 9, f"Week {week} by region", SECTION_FONT, _fill(WHITE), LEFT)
    for offset, header in enumerate(["Region", "Revenue (VND)", "Share"]):
        _paint(ws, 9, 7 + offset, header, HEADER_FONT, HEADER_FILL, CENTER, THIN)

    for offset, region_row in enumerate(regions.itertuples(index=False)):
        row = 10 + offset
        zebra = _fill(ZEBRA) if offset % 2 else _fill(WHITE)
        _paint(ws, row, 7, region_row.region, BODY_FONT, zebra, LEFT, THIN)
        _paint(ws, row, 8, int(region_row.revenue), BODY_FONT, zebra, RIGHT, THIN, VND_FORMAT)
        _paint(ws, row, 9, float(region_row.share), BODY_FONT, zebra, CENTER, THIN, SHARE_FORMAT)

    total_row = 10 + len(regions)
    _paint(ws, total_row, 7, "Total", TOTAL_FONT, _fill(TOTAL_FILL), LEFT, THIN)
    _paint(ws, total_row, 8, int(regions["revenue"].sum()), TOTAL_FONT, _fill(TOTAL_FILL), RIGHT, THIN, VND_FORMAT)
    _paint(ws, total_row, 9, 1 if len(regions) else 0, TOTAL_FONT, _fill(TOTAL_FILL), CENTER, THIN, SHARE_FORMAT)

    leader = regions.loc[regions["revenue"].idxmax()]
    insight = (
        f"{leader['region']} led week {week} with {int(leader['revenue']):,} VND "
        f"({float(leader['share']):.1%} of this week's revenue). "
        "Green week-over-week figures are increases. Red figures are decreases."
    )
    insight_row = max(14, total_row + 2)
    _merge(ws, insight_row, 2, insight_row, 9, insight, _font(11, color=INK), _fill("FFF8E7"), LEFT)
    ws.row_dimensions[insight_row].height = 32
    ws.row_dimensions[insight_row + 1].height = 12
    # With the three sample regions the yellow note is row 15, so this starts at row 17.
    _write_commentary(ws, sales, week, insight_row + 2)
    ws.row_dimensions[8].height = 22
    ws.print_title_rows = "1:2"
    ws.freeze_panes = "A4"


def write_pivot_sheet(wb: Workbook, title: str, pivot: pd.DataFrame, heading: str, tab_color: str) -> None:
    """Write a pivot table with a total row and a total column."""
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab_color
    _prepare_page(ws)

    last_column = 1 + len(pivot.columns)
    _merge(ws, 1, 1, 1, last_column, heading, TITLE_FONT, _fill(NAVY), LEFT)
    _merge(
        ws,
        2,
        1,
        2,
        last_column,
        "Revenue in VND. The Total row and the Total column are sums.",
        SUBTITLE_FONT,
        _fill(NAVY_LIGHT),
        LEFT,
    )
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20

    header_row = 4
    _paint(ws, header_row, 1, pivot.index.name, HEADER_FONT, HEADER_FILL, CENTER, THIN)
    for offset, column_name in enumerate(pivot.columns, start=2):
        _paint(ws, header_row, offset, column_name, HEADER_FONT, HEADER_FILL, CENTER, THIN)

    for row_offset, (index_value, series) in enumerate(pivot.iterrows()):
        row = header_row + 1 + row_offset
        is_total_row = str(index_value) == "Total"
        fill = _fill(TOTAL_FILL) if is_total_row else _fill(ZEBRA if row_offset % 2 else WHITE)
        font = TOTAL_FONT if is_total_row else BODY_FONT
        _paint(ws, row, 1, index_value, font, fill, LEFT, THIN)
        for column_offset, value in enumerate(series.tolist(), start=2):
            number_font = TOTAL_FONT if is_total_row or pivot.columns[column_offset - 2] == "Total" else font
            _paint(ws, row, column_offset, _as_int(value), number_font, fill, RIGHT, THIN, VND_FORMAT)
        ws.row_dimensions[row].height = 18

    ws.freeze_panes = "B5"
    ws.print_title_rows = "1:4"
    autosize_columns(ws, ignore_rows={1, 2}, min_width=12, max_width=24)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 22)


def write_trend_sheet(wb: Workbook, trend: pd.DataFrame, products: pd.DataFrame, week: int) -> None:
    """Weekly revenue table plus a line chart, and a top-products bar chart."""
    ws = wb.create_sheet("Weekly Trend")
    ws.sheet_properties.tabColor = "C65911"
    _prepare_page(ws)

    _merge(ws, 1, 1, 1, 4, "Revenue by week", SECTION_FONT, _fill(WHITE), LEFT)
    headers = ["Week", "Revenue (VND)", "Orders"]
    for offset, header in enumerate(headers, start=1):
        _paint(ws, 2, offset, header, HEADER_FONT, HEADER_FILL, CENTER, THIN)

    first_data_row = 3
    for offset, trend_row in enumerate(trend.itertuples(index=False)):
        row = first_data_row + offset
        zebra = _fill(ZEBRA) if offset % 2 else _fill(WHITE)
        _paint(ws, row, 1, int(trend_row.week), BODY_FONT, zebra, CENTER, THIN)
        _paint(ws, row, 2, int(trend_row.revenue), BODY_FONT, zebra, RIGHT, THIN, VND_FORMAT)
        _paint(ws, row, 3, int(trend_row.orders), BODY_FONT, zebra, RIGHT, THIN, VND_FORMAT)
    last_trend_row = first_data_row + len(trend) - 1

    line = LineChart()
    line.title = "Weekly revenue"
    line.style = 10
    line.y_axis.title = "Revenue (VND)"
    line.x_axis.title = "Week"
    line.y_axis.numFmt = VND_FORMAT
    line.y_axis.scaling.min = 0
    line.height = 7
    line.width = 18
    line.legend = None
    revenue_ref = Reference(ws, min_col=2, min_row=2, max_row=last_trend_row)
    week_ref = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_trend_row)
    line.add_data(revenue_ref, titles_from_data=True)
    line.set_categories(week_ref)
    line.series[0].smooth = True
    line.series[0].marker = Marker(symbol="circle", size=7)
    line.series[0].graphicalProperties.line.solidFill = NAVY
    line.series[0].graphicalProperties.line.width = 25000
    ws.add_chart(line, "E2")

    # Leave room under the line chart before the bar chart starts.
    top_title_row = max(last_trend_row + 3, 18)
    product_count = len(products)
    top_heading = f"Top {product_count} products in week {week}"
    _merge(ws, top_title_row, 1, top_title_row, 4, top_heading, SECTION_FONT, _fill(WHITE), LEFT)
    top_header_row = top_title_row + 1
    for offset, header in enumerate(["Rank", "Product", "Revenue (VND)", "Quantity"], start=1):
        _paint(ws, top_header_row, offset, header, HEADER_FONT, HEADER_FILL, CENTER, THIN)

    top_first_row = top_header_row + 1
    for offset, product_row in enumerate(products.itertuples(index=False)):
        row = top_first_row + offset
        zebra = _fill(ZEBRA) if offset % 2 else _fill(WHITE)
        _paint(ws, row, 1, int(product_row.rank), BODY_FONT, zebra, CENTER, THIN)
        _paint(ws, row, 2, product_row.product, BODY_FONT, zebra, LEFT, THIN)
        _paint(ws, row, 3, int(product_row.revenue), BODY_FONT, zebra, RIGHT, THIN, VND_FORMAT)
        _paint(ws, row, 4, int(product_row.quantity), BODY_FONT, zebra, RIGHT, THIN, VND_FORMAT)
    top_last_row = top_first_row + product_count - 1

    bars = BarChart()
    bars.type = "bar"
    bars.style = 10
    bars.title = top_heading
    bars.x_axis.title = "Revenue (VND)"
    bars.x_axis.numFmt = VND_FORMAT
    # Excel draws the first category at the bottom. Flip the axis so rank 1 is on top.
    bars.y_axis.scaling.orientation = "maxMin"
    bars.legend = None
    bars.height = 10
    bars.width = 18
    product_revenue = Reference(ws, min_col=3, min_row=top_header_row, max_row=top_last_row)
    product_names = Reference(ws, min_col=2, min_row=top_first_row, max_row=top_last_row)
    bars.add_data(product_revenue, titles_from_data=True)
    bars.set_categories(product_names)
    ws.add_chart(bars, f"E{top_title_row}")

    autosize_columns(ws, ignore_rows={1, top_title_row}, min_width=12, max_width=32)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 24)


def write_raw_sheet(wb: Workbook, sales: pd.DataFrame) -> None:
    """Write every source row as an Excel Table with filters and a frozen header."""
    ws = wb.create_sheet("Raw Data")
    ws.sheet_properties.tabColor = "7F7F7F"
    _prepare_page(ws, fit_width=True)

    ordered = sales.sort_values(["date", "region", "store", "product"])
    headers = list(COLUMN_LABELS.values())
    for column, header in enumerate(headers, start=1):
        _paint(ws, 1, column, header, HEADER_FONT, HEADER_FILL, CENTER, THIN)
    ws.row_dimensions[1].height = 22

    number_formats = {
        "date": DATE_FORMAT,
        "week": "0",
        "quantity": VND_FORMAT,
        "unit_price": VND_FORMAT,
        "revenue": VND_FORMAT,
    }
    source_columns = list(COLUMN_LABELS)
    for row_offset, record in enumerate(ordered.itertuples(index=False), start=2):
        values = {column: getattr(record, column) for column in source_columns}
        for column_offset, column_name in enumerate(source_columns, start=1):
            value = values[column_name]
            if column_name == "date":
                value = pd.Timestamp(value).date()
            elif column_name in ("week", "quantity", "unit_price", "revenue"):
                value = int(value)
            # No fill here: a cell color would hide the table's striped rows.
            cell = _paint(
                ws,
                row_offset,
                column_offset,
                value,
                BODY_FONT,
                None,
                RIGHT if column_name in number_formats and column_name != "date" else LEFT,
                THIN,
                number_formats.get(column_name),
            )
            if column_name == "date":
                cell.alignment = CENTER
            if column_name == "week":
                cell.alignment = CENTER

    last_row = 1 + len(ordered)
    last_column = get_column_letter(len(headers))
    table = Table(displayName="SalesData", ref=f"A1:{last_column}{last_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)
    ws.freeze_panes = "A2"
    ws.print_title_rows = "1:1"
    autosize_columns(ws, min_width=12, max_width=36)


def build_report(csv_path: str | Path, week: int, output_path: str | Path) -> Path:
    """Create the weekly workbook and return the path it was saved to.

    Sheets
    ------
    Summary
        KPI cards for `week`, compared with the previous week, plus bilingual commentary.
    By Region, By Category
        Revenue pivots for week 1 through `week`, with totals.
    Weekly Trend
        Line chart of weekly revenue and a bar chart of the top 10 products.
    Raw Data
        The rows those sheets were calculated from.
    """
    week = int(week)
    if week < 1:
        raise ValueError("week must be 1 or greater.")

    sales = load_sales(csv_path)
    available = sorted(int(value) for value in sales["week"].unique())
    if week not in available:
        raise ValueError(
            f"Week {week} is not in the file. Found weeks {available[0]} to {available[-1]}."
        )

    scoped = rows_through_week(sales, week)
    kpis = compute_kpis(scoped, week)
    regions = revenue_by_region(scoped, week)
    by_region = pivot_revenue(scoped, "region", "Region", sort_regions(scoped["region"].unique()))
    category_order = (
        scoped.groupby("product_category")["revenue"].sum().sort_values(ascending=False).index.tolist()
    )
    by_category = pivot_revenue(scoped, "product_category", "Category", category_order)
    trend = weekly_totals(scoped)
    products = top_products(scoped, week, limit=10)

    workbook = Workbook()
    write_summary(workbook.active, scoped, kpis, regions)
    write_pivot_sheet(workbook, "By Region", by_region, "Revenue by region and week", "2E75B6")
    write_pivot_sheet(workbook, "By Category", by_category, "Revenue by category and week", "548235")
    write_trend_sheet(workbook, trend, products, week)
    write_raw_sheet(workbook, scoped)
    workbook.active = workbook["Summary"]
    workbook.properties.title = f"Weekly Sales Report — Week {week}"
    workbook.properties.creator = "excel-report-automation"

    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    return destination
