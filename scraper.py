#!/usr/bin/env python3
"""
Multi-Page Web Scraper: Exports Product Data to Excel
=====================================================

Scrapes every product from books.toscrape.com (a site built for scraping
practice), then exports a formatted Excel workbook. Each run is compared
against the previous run, and any price changes are highlighted.

Features
--------
* Follows pagination automatically
* Retries failed requests with backoff (handles 429/5xx errors)
* Polite delay between requests
* Logs everything to a file and the console instead of crashing
* Price-change detection between runs
* Excel output: Products, Price Changes and Summary sheets

Usage
-----
    python scraper.py                    # scrape all pages
    python scraper.py --max-pages 3      # quick test run
    python scraper.py --output my.xlsx   # custom output file
"""

import argparse
import json
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

START_URL = "https://books.toscrape.com/catalogue/page-1.html"
HISTORY_FILE = Path("data/price_history.json")
LOG_FILE = Path("scraper.log")

RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("scraper")


# --------------------------------------------------------------------------
# Networking
# --------------------------------------------------------------------------
def build_session() -> requests.Session:
    """Create a requests session that retries on temporary server errors."""
    session = requests.Session()
    retry = Retry(
        total=4,
        backoff_factor=1,  # waits 1s, 2s, 4s, 8s between attempts
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {"User-Agent": "Mozilla/5.0 (portfolio sample scraper; educational use)"}
    )
    return session


def fetch(session: requests.Session, url: str) -> str:
    """Download a page and return its HTML. Raises on failure."""
    response = session.get(url, timeout=15)
    response.raise_for_status()
    response.encoding = "utf-8"  # prevents the £ sign turning into "Â£"
    return response.text


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------
def parse_page(html: str, page_url: str):
    """Extract all products on one page and find the 'next page' URL."""
    soup = BeautifulSoup(html, "html.parser")
    products = []

    for pod in soup.select("article.product_pod"):
        try:
            link = pod.select_one("h3 a")
            title = link["title"].strip()
            url = urljoin(page_url, link["href"])

            price_text = pod.select_one("p.price_color").get_text(strip=True)
            price = float(re.sub(r"[^\d.]", "", price_text))

            rating_classes = pod.select_one("p.star-rating")["class"]
            rating_word = next((c for c in rating_classes if c != "star-rating"), "")
            rating = RATING_MAP.get(rating_word)

            availability = pod.select_one("p.availability").get_text(strip=True)

            products.append(
                {
                    "Title": title,
                    "Price": price,
                    "Rating": rating,
                    "Availability": availability,
                    "URL": url,
                }
            )
        except (AttributeError, KeyError, TypeError, ValueError) as exc:
            # One malformed product should never stop the whole run.
            log.warning("Skipped a product on %s (%s)", page_url, exc)

    next_link = soup.select_one("li.next > a")
    next_url = urljoin(page_url, next_link["href"]) if next_link else None
    return products, next_url


def scrape_all(max_pages=None, delay=0.5, fetcher=fetch):
    """Walk through every page and return a list of product dicts."""
    session = build_session()
    all_products = []
    url = START_URL
    page_number = 0

    while url:
        page_number += 1
        if max_pages and page_number > max_pages:
            break

        try:
            html = fetcher(session, url)
        except requests.RequestException as exc:
            log.error("Giving up on page %d (%s): %s", page_number, url, exc)
            log.error("Saving the %d products collected so far.", len(all_products))
            break

        products, url = parse_page(html, url)
        all_products.extend(products)
        log.info("Page %d: %d products (total %d)", page_number, len(products), len(all_products))

        if url:
            time.sleep(delay)

    return all_products, page_number if not url else page_number - 1


# --------------------------------------------------------------------------
# Price history / change detection
# --------------------------------------------------------------------------
def load_history(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        log.warning("Could not read price history (%s). Starting fresh.", exc)
        return {}


def save_history(path: Path, products: list) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    history = {p["URL"]: p["Price"] for p in products}
    path.write_text(json.dumps(history, indent=2), encoding="utf-8")


def build_dataframe(products: list, history: dict, first_run: bool = False) -> pd.DataFrame:
    df = pd.DataFrame(products)
    df["Previous Price"] = df["URL"].map(history)  # NaN where not seen before
    df["Change"] = (df["Price"] - df["Previous Price"]).round(2)

    def status(row) -> str:
        if first_run:
            return "First run"
        if pd.isna(row["Previous Price"]):
            return "New"
        return "Changed" if row["Change"] != 0 else "Unchanged"

    df["Status"] = df.apply(status, axis=1)
    df = df.rename(
        columns={"Price": "Price (GBP)", "Rating": "Rating (1-5)"}
    )
    return df[
        ["Title", "Price (GBP)", "Previous Price", "Change", "Status",
         "Rating (1-5)", "Availability", "URL"]
    ]


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
CHANGED_FILL = PatternFill("solid", fgColor="FFF2CC")


def style_sheet(ws, widths: dict, money_cols=(), highlight_changes=False):
    """Apply header style, widths, filters and number formats to a sheet."""
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for col_idx in money_cols:
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            row[0].number_format = '"£"#,##0.00'

    if highlight_changes:
        for row in ws.iter_rows(min_row=2):
            change = row[3].value  # "Change" column
            if isinstance(change, (int, float)) and change != 0:
                for cell in row:
                    cell.fill = CHANGED_FILL

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def export_excel(df: pd.DataFrame, output: Path, pages: int, first_run: bool) -> None:
    changes = df[df["Change"].notna() & (df["Change"] != 0)]

    summary = pd.DataFrame(
        {
            "Metric": [
                "Scrape date",
                "Pages scraped",
                "Total products",
                "Average price (GBP)",
                "Cheapest price (GBP)",
                "Most expensive price (GBP)",
                "Average rating",
                "Price changes since last run",
                "New products since last run",
            ],
            "Value": [
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                pages,
                len(df),
                round(df["Price (GBP)"].mean(), 2),
                df["Price (GBP)"].min(),
                df["Price (GBP)"].max(),
                round(df["Rating (1-5)"].mean(), 2),
                "n/a (first run)" if first_run else len(changes),
                "n/a (first run)" if first_run else int((df["Status"] == "New").sum()),
            ],
        }
    )

    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Products", index=False)
        changes.to_excel(writer, sheet_name="Price Changes", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        widths = {1: 60, 2: 14, 3: 15, 4: 10, 5: 12, 6: 14, 7: 15, 8: 70}
        style_sheet(writer.sheets["Products"], widths, money_cols=(2, 3, 4),
                    highlight_changes=True)
        style_sheet(writer.sheets["Price Changes"], widths, money_cols=(2, 3, 4))
        style_sheet(writer.sheets["Summary"], {1: 32, 2: 22})

        if changes.empty:
            writer.sheets["Price Changes"]["A2"] = (
                "No price changes since the last run."
            )

        money_labels = {"Average price (GBP)", "Cheapest price (GBP)",
                        "Most expensive price (GBP)"}
        for row in writer.sheets["Summary"].iter_rows(min_row=2):
            label, value_cell = row[0].value, row[1]
            value_cell.alignment = Alignment(horizontal="left")
            if label in money_labels:
                value_cell.number_format = '"£"#,##0.00'
            elif label == "Average rating":
                value_cell.number_format = "0.00"


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def setup_logging() -> None:
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", "%H:%M:%S")
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(fmt)
    log.setLevel(logging.INFO)
    log.addHandler(console)
    log.addHandler(file_handler)


def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape books.toscrape.com to Excel")
    parser.add_argument("--max-pages", type=int, default=None,
                        help="stop after this many pages (default: all 50)")
    parser.add_argument("--delay", type=float, default=0.5,
                        help="seconds to wait between pages (default: 0.5)")
    parser.add_argument("--output", type=Path, default=Path("output/books.xlsx"),
                        help="Excel file to create (default: output/books.xlsx)")
    args = parser.parse_args()

    setup_logging()
    started = time.time()
    log.info("Starting scrape...")

    products, pages = scrape_all(args.max_pages, args.delay)
    if not products:
        log.error("No products collected. Nothing to export.")
        return 1

    first_run = not HISTORY_FILE.exists()
    history = load_history(HISTORY_FILE)
    df = build_dataframe(products, history, first_run)

    export_excel(df, args.output, pages, first_run)
    save_history(HISTORY_FILE, products)

    changed = df["Change"].notna() & (df["Change"] != 0)
    elapsed = time.time() - started
    log.info("Done: %d products from %d pages in %.1f seconds.", len(df), pages, elapsed)
    if not first_run:
        log.info("Price changes since last run: %d", int(changed.sum()))
    log.info("Excel file saved to: %s", args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
