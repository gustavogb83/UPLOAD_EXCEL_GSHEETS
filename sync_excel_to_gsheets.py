#!/usr/bin/env python3
"""Sync daily orders/revenue from the CLA Excel workbook to one Google Sheet per market.

Run daily (cron / Task Scheduler). Only the last `lookback_days` days are written, matched by date
on the Daily tab of each market's sheet. Safe to re-run: it only overwrites the order/revenue cells.
"""
import argparse
import hashlib
import json
import logging
import re
import shutil
import sys
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

log = logging.getLogger("sync")

MONTHS = {m: i + 1 for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"])}
TAB_RE = re.compile(r"^CLA\s+([A-Z]{3,9})\s+(\d{2})\s+NEW\s*$", re.I)
MARKET_LABELS = {"SE": "sweden", "NO": "norway", "DK": "denmark", "FI": "finland"}
CURRENCY = {"SE": "sek", "NO": "nok", "DK": "dkk", "FI": "€"}
TEXT_DATE_FORMATS = ["%A, %B %d, %Y", "%Y-%m-%d", "%d/%m/%Y"]
FR_MONTHS = {"janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
             "juillet": 7, "août": 8, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11,
             "décembre": 12, "decembre": 12}


def to_date(v):
    """Convert an Excel cell (datetime or text) to a date, or None."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        s = v.strip().replace("\xa0", " ")
        for f in TEXT_DATE_FORMATS:
            try:
                return datetime.strptime(s, f).date()
            except ValueError:
                pass
        m = re.match(r"^\w+\s+(\d{1,2})\s+(\w+)\s+(\d{4})$", s)  # "jeudi 1 janvier 2026"
        if m and m.group(2).lower() in FR_MONTHS:
            return date(int(m.group(3)), FR_MONTHS[m.group(2).lower()], int(m.group(1)))
    return None


def row_date(ws, r):
    """Real date cells win over text (Feb's column A holds stale text like 'jeudi 1 janvier 2026')."""
    vals = [ws.cell(r, 1).value, ws.cell(r, 2).value]
    for v in vals:
        if isinstance(v, (datetime, date)):
            return to_date(v)
    return to_date(vals[0]) or to_date(vals[1])


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def tab_month(name):
    m = TAB_RE.match(name)
    if not m or m.group(1)[:3].upper() not in MONTHS:
        return None
    return 2000 + int(m.group(2)), MONTHS[m.group(1)[:3].upper()]


def find_columns(ws):
    """Locate revenue/orders columns per market from the header rows (row 2 = market, row 3 = field).
    Column positions differ between monthly tabs, so they are detected, not hard-coded."""
    labels = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(2, c).value
        if isinstance(v, str):
            labels[c] = v.strip().lower()
    order = sorted(labels)
    cols = {}
    for mk, label in MARKET_LABELS.items():
        start = next((c for c in order if labels[c] == label), None)
        if start is None:
            raise ValueError(f"market '{label}' not found in row 2 of '{ws.title}'")
        nxt = next((c for c in order if c > start), ws.max_column + 1)
        rev = orders = None
        for c in range(start, nxt):
            h = str(ws.cell(3, c).value or "").lower().replace("\n", " ")
            if "real" in h and "revenue" in h and CURRENCY[mk] in h:
                rev = c
            elif "real" in h and ("order" in h or "oder" in h):
                orders = c
        if rev is None or orders is None:
            raise ValueError(f"{mk}: revenue/orders columns not found in '{ws.title}'")
        cols[mk] = (rev, orders)
    return cols


def read_excel(path, start, end):
    """Return {market: {date: (revenue, orders)}} for dates in [start, end]."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    out = {mk: {} for mk in MARKET_LABELS}
    for ws in wb.worksheets:
        ym = tab_month(ws.title)
        if not ym or ws.sheet_state != "visible":
            continue
        if (ym[0], ym[1]) < (start.year, start.month) or (ym[0], ym[1]) > (end.year, end.month):
            continue
        cols = find_columns(ws)
        for r in range(4, ws.max_row + 1):
            d = row_date(ws, r)
            if d is None or not (start <= d <= end) or (d.year, d.month) != ym:
                continue
            for mk, (rc, oc) in cols.items():
                rev, orders = num(ws.cell(r, rc).value), num(ws.cell(r, oc).value)
                if rev is None and orders is None:
                    continue  # day not filled in yet
                out[mk][d] = (rev, orders)
    return out


def sheet_date(v):
    if isinstance(v, (int, float)):  # Google serial number
        return date(1899, 12, 30) + timedelta(days=int(v))
    if isinstance(v, str):
        s = v.strip()
        for f in ("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(s, f).date()
            except ValueError:
                pass
    return None


def push_market(cfg, mk, data, dry_run):
    import gspread
    gc = gspread.service_account(filename=cfg["service_account_file"])
    ws = gc.open_by_key(cfg["markets"][mk]).worksheet(cfg.get("sheet_tab", "Daily"))
    dcol, ocol, rcol = (cfg["sheet_columns"][k] for k in ("date", "orders", "revenue"))
    dates = ws.get(f"{dcol}:{dcol}", value_render_option="UNFORMATTED_VALUE")
    row_of = {}
    for i, cell in enumerate(dates, start=1):
        d = sheet_date(cell[0]) if cell else None
        if d:
            row_of.setdefault(d, i)
    updates, missing = [], []
    for d, (rev, orders) in sorted(data.items()):
        row = row_of.get(d)
        if row is None:
            missing.append(d)
            continue
        if orders is not None:
            updates.append({"range": f"{ocol}{row}", "values": [[orders]]})
        if rev is not None:
            updates.append({"range": f"{rcol}{row}", "values": [[rev]]})
    log.info("%s: %d cells to write, %d dates not found in sheet %s", mk, len(updates), len(missing),
             [d.isoformat() for d in missing] or "")
    if updates and not dry_run:
        ws.batch_update(updates, value_input_option="RAW")
    elif dry_run:
        for u in updates:
            log.info("  [dry-run] %s %s = %s", mk, u["range"], u["values"][0][0])
    return len(missing)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--config", default="config.json")
    p.add_argument("--dry-run", action="store_true", help="print what would be written")
    p.add_argument("--force", action="store_true", help="sync even if the file is unchanged")
    p.add_argument("--from", dest="from_", help="backfill start date YYYY-MM-DD (default: today - lookback)")
    p.add_argument("--to", help="end date YYYY-MM-DD (default: today)")
    a = p.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    cfg = json.loads(Path(a.config).read_text())
    end = datetime.strptime(a.to, "%Y-%m-%d").date() if a.to else date.today()
    start = (datetime.strptime(a.from_, "%Y-%m-%d").date() if a.from_
             else end - timedelta(days=cfg.get("lookback_days", 15)))

    src = Path(cfg["excel_path"])
    state_file = Path(cfg.get("state_file", "sync_state.json"))
    digest = hashlib.sha256(src.read_bytes()).hexdigest()
    state = json.loads(state_file.read_text()) if state_file.exists() else {}
    if not a.force and not a.dry_run and not a.from_ and state.get("sha256") == digest \
            and state.get("day") == date.today().isoformat():
        log.info("File unchanged since today's last sync; nothing to do (use --force to override)")
        return 0

    with tempfile.TemporaryDirectory() as td:  # work on a copy: the file may be open in Excel
        tmp = Path(td) / "snapshot.xlsx"
        shutil.copy2(src, tmp)
        data = read_excel(tmp, start, end)
    log.info("Window %s..%s: %s", start, end, {mk: len(v) for mk, v in data.items()})

    failures = 0
    for mk in MARKET_LABELS:
        try:
            push_market(cfg, mk, data[mk], a.dry_run)
        except Exception:
            log.exception("%s failed", mk)
            failures += 1
    if not failures and not a.dry_run:
        state_file.write_text(json.dumps({"sha256": digest, "day": date.today().isoformat()}))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
