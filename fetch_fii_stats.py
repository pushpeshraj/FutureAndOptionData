"""
Fetch the latest available NSE 'FII Derivative Statistics' report and write it as
fii_stats.json (+ raw file) for the dashboard to display.

Uses ONLY `requests` + the Python standard library (no pandas/xlrd needed).
NSE serves this report as a binary Excel (.xls) file in which every cell — including
the numbers — is stored as a shared string, so we read the OLE2 container and the
shared-string table directly. An HTML-table fallback is included in case NSE ever
switches formats.

Run by the GitHub Action, or locally:
    python fetch_fii_stats.py
"""

import io
import re
import json
import struct
import datetime as dt
import zoneinfo
from html.parser import HTMLParser

import requests

# dd = zero-padded day, mon = English 3-letter month, yyyy = 4-digit year.
PATH = "/content/fo/fii_stats_{dd}-{mon}-{yyyy}.xls"
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


# --------------------------------------------------------------------------- #
# Pure-stdlib readers
# --------------------------------------------------------------------------- #
def _xls_to_matrix(data):
    """Read a binary (OLE2/BIFF8) .xls whose cells are all shared strings."""
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError("not an OLE2 .xls")
    ssz = 1 << struct.unpack_from("<H", data, 30)[0]   # sector size
    msz = 1 << struct.unpack_from("<H", data, 32)[0]   # mini-sector size
    dir_start = struct.unpack_from("<I", data, 48)[0]
    mini_cutoff = struct.unpack_from("<I", data, 56)[0]
    minifat_start = struct.unpack_from("<I", data, 60)[0]
    EOC, FREE = 0xFFFFFFFE, 0xFFFFFFFF

    def soff(sid):
        return (sid + 1) * ssz

    difat = list(struct.unpack_from("<109i", data, 76))
    fat = []
    for fsid in difat:
        if fsid < 0:
            continue
        fat.extend(struct.unpack_from("<%dI" % (ssz // 4), data, soff(fsid)))

    def chain(start, size=None):
        out, sid = b"", start
        while sid not in (EOC, FREE) and sid < len(fat):
            out += data[soff(sid):soff(sid) + ssz]
            sid = fat[sid]
        return out[:size] if size else out

    dir_data = chain(dir_start)
    root = wb = None
    for i in range(0, len(dir_data), 128):
        e = dir_data[i:i + 128]
        if len(e) < 128:
            break
        nl = struct.unpack_from("<H", e, 64)[0]
        if nl == 0:
            continue
        name = e[:nl - 2].decode("utf-16-le", "ignore")
        typ = e[66]
        ssid = struct.unpack_from("<I", e, 116)[0]
        sz = struct.unpack_from("<I", e, 120)[0]
        if typ == 5:
            root = (ssid, sz)
        if name == "Workbook" or name == "Book":
            wb = (ssid, sz)
    if wb is None:
        raise ValueError("no Workbook stream found")

    if wb[1] >= mini_cutoff:
        stream = chain(wb[0], wb[1])
    else:                                              # small stream -> mini FAT
        mini = chain(root[0], root[1])
        mf = chain(minifat_start)
        minifat = list(struct.unpack_from("<%dI" % (len(mf) // 4), mf, 0))
        out, sid = b"", wb[0]
        while sid not in (EOC, FREE) and sid < len(minifat):
            out += mini[sid * msz:sid * msz + msz]
            sid = minifat[sid]
        stream = out[:wb[1]]

    # Collect records; fold CONTINUE (0x003C) payloads into the previous record.
    recs, pos = [], 0
    while pos + 4 <= len(stream):
        rt, rl = struct.unpack_from("<HH", stream, pos)
        pos += 4
        payload = stream[pos:pos + rl]
        pos += rl
        if rt == 0x003C and recs:
            recs[-1] = (recs[-1][0], recs[-1][1] + payload)
        else:
            recs.append((rt, payload))

    # Shared-string table (record 0x00FC).
    sst = []
    for rt, d in recs:
        if rt != 0x00FC:
            continue
        n = struct.unpack_from("<I", d, 4)[0]
        p = 8
        for _ in range(n):
            if p + 3 > len(d):
                break
            cch = struct.unpack_from("<H", d, p)[0]
            flags = d[p + 2]
            p += 3
            high, ext, rich = flags & 0x01, flags & 0x04, flags & 0x08
            crun = struct.unpack_from("<H", d, p)[0] if rich else 0
            if rich:
                p += 2
            cbext = struct.unpack_from("<I", d, p)[0] if ext else 0
            if ext:
                p += 4
            if high:
                s = d[p:p + cch * 2].decode("utf-16-le", "ignore"); p += cch * 2
            else:
                s = d[p:p + cch].decode("latin-1", "ignore"); p += cch
            p += crun * 4 + cbext
            sst.append(s)
        break

    # String cells (record 0x00FD = LABELSST).
    grid = {}
    for rt, d in recs:
        if rt == 0x00FD and len(d) >= 10:
            row, col = struct.unpack_from("<HH", d, 0)
            isst = struct.unpack_from("<I", d, 6)[0]
            grid[(row, col)] = sst[isst] if isst < len(sst) else ""
    if not grid:
        return []
    maxr = max(r for r, _ in grid)
    maxc = max(c for _, c in grid)
    return [[grid.get((r, c), "") for c in range(maxc + 1)] for r in range(maxr + 1)]


class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell, self._in = [], None, [], False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th"):
            self._in, self._cell = True, []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._row is not None:
            self._in = False
            self._row.append("".join(self._cell).strip())
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._in:
            self._cell.append(data)


def _read_tables(content, text):
    """Return (list_of_matrices, errors). Tries binary .xls, then HTML."""
    errors = []
    if content[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        try:
            m = _xls_to_matrix(content)
            if m:
                return [m], errors
        except Exception as e:
            errors.append(f"xls: {type(e).__name__}: {e}")
    try:
        p = _TableParser()
        p.feed(text or content.decode("latin-1", "ignore"))
        if p.rows:
            return [p.rows], errors
    except Exception as e:
        errors.append(f"html: {type(e).__name__}: {e}")
    return [], errors


# --------------------------------------------------------------------------- #
# Row extraction
# --------------------------------------------------------------------------- #
def extract_rows(matrices):
    rows, seen = [], set()
    for matrix in matrices:
        for raw in matrix:
            cells = [str(c).strip() for c in raw if str(c).strip()]
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

    mains = [r for r in rows if r["Category"] in
             ("Index Futures", "Index Options", "Stock Futures", "Stock Options")]
    if mains and not any(r["Category"] == "Total" for r in rows):
        total = {"Category": "Total"}
        for col in COLUMNS[1:]:
            s = sum(float(r[col]) for r in mains)
            total[col] = str(int(round(s))) if "Contracts" in col else f"{s:.2f}"
        rows.append(total)

    order = {c: i for i, c in enumerate(
        ["Index Futures", "Index Options", "Stock Futures", "Stock Options", "Total"])}
    rows.sort(key=lambda r: order.get(r["Category"], 99))
    return rows


def extract_report_date(matrices, fallback):
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


# --------------------------------------------------------------------------- #
# Fetch
# --------------------------------------------------------------------------- #
def fetch_latest(start_date=None, max_lookback=10):
    start_date = start_date or dt.datetime.now(IST).date()
    session = requests.Session()
    session.headers.update(HEADERS)
    session.get(HOMEPAGE, timeout=15)
    session.get(REPORTS_PAGE, timeout=15)

    saw_200, last_snippet, last_errors = False, "", []
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
            print(f"  {url} -> {resp.status_code}, {len(resp.content)} bytes")
            if resp.status_code == 200 and resp.content:
                saw_200 = True
                matrices, errors = _read_tables(resp.content, resp.text)
                rows = extract_rows(matrices)
                if rows:
                    return rows, extract_report_date(matrices, date), url, resp.content
                last_errors = errors
                last_snippet = repr(resp.content[:120])

    if saw_200:
        raise RuntimeError(
            "Downloaded the FII file but could not parse it — layout may have changed.\n"
            "Parser errors: " + "; ".join(last_errors) + "\nFirst bytes: " + last_snippet
        )
    raise FileNotFoundError(
        f"No FII stats file found (all non-200) in the {max_lookback} days before "
        f"{start_date:%Y-%m-%d}. The URL path may have changed."
    )


def main():
    rows, report_date, url, raw = fetch_latest()
    now_ist = dt.datetime.now(IST)

    with open(f"fii_stats_{report_date:%d%m%Y}.xls", "wb") as f:
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
