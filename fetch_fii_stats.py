"""
Fetch the latest available NSE 'FII Derivative Statistics' report and write it as
fii_stats.json (+ raw .xls) for the dashboard to display.

Unlike the participant-OI CSV, this report is published as an .xls file that is
actually an HTML table, so we parse it with the standard-library HTML parser
(no extra dependencies beyond requests).

Run by the GitHub Action, or locally:
    python fetch_fii_stats.py
"""

import json
import datetime as dt
import zoneinfo
import re
from html.parser import HTMLParser

import requests

# dd is zero-padded, MMM is the English 3-letter month, yyyy is 4 digits.
ARCHIVE_URL = "https://nsearchives.nseindia.com/content/fo/fii_stats_{dd}-{mon}-{yyyy}.xls"
HOMEPAGE = "https://www.nseindia.com"
REPORTS_PAGE = "https://www.nseindia.com/all-reports-derivatives"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
}

IST = zoneinfo.ZoneInfo("Asia/Kolkata")
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]  # locale-independent

# The four instrument categories plus the total row.
CATEGORIES = ["INDEX FUTURES", "INDEX OPTIONS", "STOCK FUTURES", "STOCK OPTIONS", "TOTAL"]

COLUMNS = [
    "Category",
    "Buy Contracts", "Buy Amount (Cr)",
    "Sell Contracts", "Sell Amount (Cr)",
    "OI Contracts", "OI Amount (Cr)",
]

NUMBER_RE = re.compile(r"^-?[\d,]+(\.\d+)?$")


class _TableParser(HTMLParser):
    """Collects every table row as a list of cell-text strings."""
    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell, self._in_cell = [], None, [], False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th"):
            self._in_cell, self._cell = True, []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._row is not None:
            self._in_cell = False
            self._row.append("".join(self._cell).strip())
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._in_cell:
            self._cell.append(data)


def parse_fii_html(html_text):
    """Return a list of row-dicts for the five categories, or [] if not found."""
    parser = _TableParser()
    parser.feed(html_text)

    rows = []
    for cells in parser.rows:
        if not cells:
            continue
        label = cells[0].strip().upper()
        if label in CATEGORIES:
            nums = [c.replace(",", "") for c in cells[1:] if NUMBER_RE.match(c.strip())]
            if len(nums) >= 6:
                values = [label.title()] + nums[:6]
                rows.append({COLUMNS[i]: values[i] for i in range(len(COLUMNS))})
    return rows


def fetch_latest(start_date=None, max_lookback=10):
    """Return (rows, report_date) for the most recent published FII stats report."""
    start_date = start_date or dt.datetime.now(IST).date()

    session = requests.Session()
    session.headers.update(HEADERS)
    session.get(HOMEPAGE, timeout=15)
    session.get(REPORTS_PAGE, timeout=15)

    for i in range(max_lookback):
        date = start_date - dt.timedelta(days=i)
        url = ARCHIVE_URL.format(dd=date.strftime("%d"), mon=MONTHS[date.month - 1], yyyy=date.year)
        resp = session.get(url, headers={"Referer": REPORTS_PAGE}, timeout=30)
        if resp.status_code == 200 and resp.content:
            rows = parse_fii_html(resp.text)
            if rows:                       # a real report parses into category rows
                return rows, date, url, resp.text

    raise FileNotFoundError(
        f"No FII stats report found in the {max_lookback} days before {start_date:%Y-%m-%d}."
    )


def main():
    rows, report_date, url, raw_html = fetch_latest()
    now_ist = dt.datetime.now(IST)

    # Save the raw file alongside the JSON.
    with open(f"fii_stats_{report_date:%d%m%Y}.xls", "w", encoding="utf-8") as f:
        f.write(raw_html)

    payload = {
        "title": "FII Derivative Statistics",
        "report_date": report_date.isoformat(),
        "report_date_display": report_date.strftime("%d %b %Y"),
        "fetched_at_ist": now_ist.strftime("%d %b %Y, %I:%M %p IST"),
        "source_url": url,
        "columns": COLUMNS,
        "rows": rows,
    }
    with open("fii_stats.json", "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

    print(f"Wrote fii_stats.json for report dated {report_date:%Y-%m-%d} ({len(rows)} rows).")


if __name__ == "__main__":
    main()
