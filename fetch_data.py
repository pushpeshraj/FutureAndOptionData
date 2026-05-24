"""
Fetch the latest available NSE F&O 'Participant-wise Open Interest' report
and write it out as both a raw CSV and a data.json file that the GitHub Pages
website (index.html) reads to render the table.

Run by the GitHub Action on a schedule. Can also be run locally:
    python fetch_data.py
"""

import csv
import io
import json
import datetime as dt
import zoneinfo

import requests

ARCHIVE_URL = "https://nsearchives.nseindia.com/content/nsccl/fao_participant_oi_{ddmmyyyy}.csv"
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


def fetch_latest(start_date=None, max_lookback=10):
    """Return (csv_text, report_date) for the most recent published report.

    Walks backwards day-by-day so weekends and holidays are skipped automatically.
    """
    start_date = start_date or dt.datetime.now(IST).date()

    session = requests.Session()
    session.headers.update(HEADERS)
    # Warm up once so NSE hands over its session cookies.
    session.get(HOMEPAGE, timeout=15)
    session.get(REPORTS_PAGE, timeout=15)

    for i in range(max_lookback):
        date = start_date - dt.timedelta(days=i)
        url = ARCHIVE_URL.format(ddmmyyyy=date.strftime("%d%m%Y"))
        resp = session.get(url, headers={"Referer": REPORTS_PAGE}, timeout=30)

        # Valid file = 200, has content, and isn't an HTML error page.
        if resp.status_code == 200 and resp.content and not resp.content.lstrip().startswith(b"<"):
            return resp.text, date

    raise FileNotFoundError(
        f"No report found in the {max_lookback} days before {start_date:%Y-%m-%d}."
    )


def parse_csv(csv_text):
    """Parse the NSE CSV into (title, columns, rows-as-dicts).

    Line 1 is a title row, line 2 is the column header, the rest are data.
    """
    reader = list(csv.reader(io.StringIO(csv_text)))
    # Drop fully empty lines.
    reader = [r for r in reader if any(cell.strip() for cell in r)]

    title = reader[0][0].strip() if reader and reader[0] else "Participant-wise Open Interest"
    header = [h.strip() for h in reader[1]]
    rows = []
    for raw in reader[2:]:
        # Pad/trim row to header length.
        cells = (raw + [""] * len(header))[: len(header)]
        rows.append({header[i]: cells[i].strip() for i in range(len(header))})

    return title, header, rows


def main():
    csv_text, report_date = fetch_latest()
    title, columns, rows = parse_csv(csv_text)

    now_ist = dt.datetime.now(IST)

    # Save the raw CSV alongside the JSON.
    with open(f"participant_oi_{report_date:%d%m%Y}.csv", "w", encoding="utf-8") as f:
        f.write(csv_text)

    payload = {
        "title": title,
        "report_date": report_date.isoformat(),
        "report_date_display": report_date.strftime("%d %b %Y"),
        "fetched_at_ist": now_ist.strftime("%d %b %Y, %I:%M %p IST"),
        "source_url": ARCHIVE_URL.format(ddmmyyyy=report_date.strftime("%d%m%Y")),
        "columns": columns,
        "rows": rows,
    }

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

    print(f"Wrote data.json for report dated {report_date:%Y-%m-%d} "
          f"({len(rows)} rows, {len(columns)} columns).")


if __name__ == "__main__":
    main()
