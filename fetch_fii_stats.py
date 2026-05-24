"""
Fetch the latest available NSE 'FII Derivative Statistics' report and write it as
fii_stats.json (+ raw file) for the dashboard to display.

NSE serves this report as an .xls that may be an HTML table OR a real Excel binary,
and the layout occasionally changes, so this reader tries several strategies
(read_html, then read_excel) and prints diagnostics for each attempt.

Run by the GitHub Action, or locally:
    python fetch_fii_stats.py
"""

import io
import re
import json
import datetime as dt
import zoneinfo

import requests
import pandas as pd

# dd = zero-padded day, mon = English 3-letter month, yyyy = 4-digit year.
PATH = "/content/fo/fii_stats_{dd}-{mon}-{yyyy}.xls"
# Tried in order; the first that returns the file wins.
BASES = ["https://nsearchives.nseindia.com", "https://archives.nseindia.com", "https://www1.nseindia.com"]
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
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

CATEGORIES = ["INDEX FUTURES", "INDEX OPTIONS", "STOCK FUTURES", "STOCK OPTIONS", "TOTAL"]
COLUMNS = [
    "Category",
    "Buy Contracts", "Buy Amount (Cr)",
    "Sell Contracts", "Sell Amount (Cr)",
    "OI Contracts", "OI Amount (Cr)",
]
NUMBER_RE = re.compile(r"^-?\d+(\.\d+)?$")


def _read_tables(content, text):
    """Return (list_of_matrices, list_of_error_strings). Tries HTML then Excel."""
    head = content[:4096].lower()
    looks_html = content[:1] == b"<" or b"<table" in head or b"<html" in head

    def via_html():
        dfs = pd.read_html(io.StringIO(text), header=None)
        return [df.astype(str).values.tolist() for df in dfs]

    def via_excel():
        sheets = pd.read_excel(io.BytesIO(content), header=None, sheet_name=None)
        return [df.astype(str).values.tolist() for df in sheets.values()]

    order = [("read_html", via_html), ("read_excel", via_excel)]
    if not looks_html:
        order.reverse()

    matrices, errors = [], []
    for name, fn in order:
        try:
            m = fn()
            if any(rows for rows in m):
                return m, errors
        except Exception as e:
            errors.append(f"{name}: {type(e).__name__}: {e}")
    return matrices, errors


def extract_rows(matrices):
    """Pull the 5 category rows out of whatever tables we parsed."""
    rows, seen = [], set()
    for matrix in matrices:
        for raw in matrix:
            cells = [str(c).strip() for c in raw if str(c).strip().lower() != "nan"]
            if not cells:
                continue
            label = None
            for c in cells[:2]:
                key = re.sub(r"\s+", " ", c).strip().upper()
                if key in CATEGORIES:
                    label = key
                    break
            if not label or label in seen:
                continue
            nums = []
            for c in cells:
                cc = c.replace(",", "").replace("\u20b9", "").strip()
                if cc.endswith(".0"):
                    cc = cc[:-2]
                if NUMBER_RE.match(cc):
                    nums.append(cc)
            if len(nums) >= 6:
                vals = [label.title()] + nums[:6]
                rows.append({COLUMNS[i]: vals[i] for i in range(len(COLUMNS))})
                seen.add(label)

    # The NSE file has no TOTAL row, so compute one from the four headline
    # categories (each is already the aggregate of its instrument sub-rows).
    mains = [r for r in rows if r["Category"] in
             ("Index Futures", "Index Options", "Stock Futures", "Stock Options")]
    if mains and not any(r["Category"] == "Total" for r in rows):
        total = {"Category": "Total"}
        for col in COLUMNS[1:]:
            s = sum(float(r[col]) for r in mains)
            total[col] = str(int(round(s))) if "Contracts" in col else f"{s:.2f}"
        rows.append(total)

    # Keep canonical order.
    order = {c: i for i, c in enumerate(
        ["Index Futures", "Index Options", "Stock Futures", "Stock Options", "Total"])}
    rows.sort(key=lambda r: order.get(r["Category"], 99))
    return rows


def extract_report_date(matrices, fallback):
    """Read the date from the file's own header (e.g. '... FOR 22-May-2026')."""
    for matrix in matrices:
        for raw in matrix:
            for cell in raw:
                m = re.search(r"(\d{1,2})-([A-Za-z]{3})-(\d{4})", str(cell))
                if m:
                    try:
                        return dt.datetime.strptime(m.group(0), "%d-%b-%Y").date()
                    except ValueError:
                        pass
    return fallback


def fetch_latest(start_date=None, max_lookback=10):
    start_date = start_date or dt.datetime.now(IST).date()
    session = requests.Session()
    session.headers.update(HEADERS)
    session.get(HOMEPAGE, timeout=15)
    session.get(REPORTS_PAGE, timeout=15)

    saw_200 = False
    last_snippet = ""
    last_errors = []

    for i in range(max_lookback):
        date = start_date - dt.timedelta(days=i)
        path = PATH.format(dd=date.strftime("%d"), mon=MONTHS[date.month - 1], yyyy=date.year)
        for base in BASES:
            url = base + path
            try:
                resp = session.get(url, headers={"Referer": REPORTS_PAGE}, timeout=30)
            except requests.RequestException as e:
                print(f"  {url} -> ERROR {e}")
                continue
            ctype = resp.headers.get("Content-Type", "?")
            print(f"  {url} -> {resp.status_code}, {len(resp.content)} bytes, {ctype}")
            if resp.status_code == 200 and resp.content:
                saw_200 = True
                matrices, errors = _read_tables(resp.content, resp.text)
                rows = extract_rows(matrices)
                if rows:
                    report_date = extract_report_date(matrices, date)
                    return rows, report_date, url, resp.text
                last_errors = errors
                last_snippet = resp.text[:400] if resp.text else repr(resp.content[:120])

    if saw_200:
        raise RuntimeError(
            "Downloaded the FII file but could not parse category rows — the layout "
            "may have changed.\nParser errors: " + "; ".join(last_errors) +
            "\nFirst bytes of last response:\n" + last_snippet
        )
    raise FileNotFoundError(
        f"No FII stats file found (all requests non-200) in the {max_lookback} days "
        f"before {start_date:%Y-%m-%d}. The URL path may have changed."
    )


def main():
    rows, report_date, url, raw = fetch_latest()
    now_ist = dt.datetime.now(IST)

    with open(f"fii_stats_{report_date:%d%m%Y}.xls", "w", encoding="utf-8") as f:
        f.write(raw)

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

    print(f"Wrote fii_stats.json for {report_date:%Y-%m-%d} ({len(rows)} rows) from {url}")


if __name__ == "__main__":
    main()
