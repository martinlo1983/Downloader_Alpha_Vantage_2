"""
Alpha Vantage → Google Sheets fetcher
Sheet: Datos_Alpha_Vantage (1bzBl-2NuCf8Iwu3_YC1XZYM7UQ63m3YQGVJKlaDGRxg)
  - Input:  pestaña TICKERS  (col A: TICKERS, fila 1 = header)
  - Output: pestaña OUT       (7 filas por ticker)
  - State:  pestaña AV_STATUS (creada automáticamente)

3 calls/ticker | 23 req/day | 7 tickers/day | ~43-day cycle for 300 tickers

Output rows per ticker:
  PER         Q(n)    precio / EPS quarter más reciente
  PER         Q(n-1)
  PER         Q(n-2)
  PER         Q(n-3)
  PER         Q(n-4)
  PER_HIST_3A Inst    promedio PER últimos 12 quarters (3 años)
  PEG 5A      Q(n)    PEGRatio de OVERVIEW
"""

import os, time, json, logging
from datetime import datetime, date
from typing import Optional

import requests
import gspread
from google.oauth2.service_account import Credentials

# ── Config ────────────────────────────────────────────────────────────────────
AV_API_KEY        = os.environ["AV_API_KEY"]
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]
SHEET_ID          = os.environ.get("SHEET_ID", "1bzBl-2NuCf8Iwu3_YC1XZYM7UQ63m3YQGVJKlaDGRxg")

TICKER_SHEET      = "TICKERS"
OUT_SHEET         = "OUT"
STATUS_SHEET      = "AV_STATUS"

CALLS_PER_TICKER  = 3
REQUESTS_PER_DAY  = 23
TICKERS_PER_DAY   = REQUESTS_PER_DAY // CALLS_PER_TICKER   # = 7
UPDATE_CYCLE_DAYS = 45
AV_DELAY_SECS     = 13

OUT_HEADER = ["FECHA CARGA INFO", "ACCION", "DATO", "TEMP", "VALOR"]

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)


# ── Google Sheets ─────────────────────────────────────────────────────────────
def open_workbook():
    creds_info = json.loads(GOOGLE_CREDS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds).open_by_key(SHEET_ID)


def ensure_sheet(wb, name: str, rows=400, cols=5) -> gspread.Worksheet:
    try:
        return wb.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = wb.add_worksheet(name, rows=rows, cols=cols)
        log.info("Created tab: %s", name)
        return ws


def ensure_out_header(out_ws):
    """Write header row if OUT sheet is empty."""
    vals = out_ws.get_all_values()
    if not vals or vals[0] != OUT_HEADER:
        out_ws.clear()
        out_ws.update("A1", [OUT_HEADER])


# ── Status sheet ──────────────────────────────────────────────────────────────
def load_status(status_ws) -> dict:
    # Use get_all_values() instead of get_all_records() to avoid
    # GSpreadException when header row has duplicate or missing columns.
    all_rows = status_ws.get_all_values()
    out = {}
    for row in all_rows[1:]:   # skip header
        if not row or not row[0].strip():
            continue
        t   = row[0].strip().upper()
        raw = row[1] if len(row) > 1 else ""
        try:
            last = datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            last = date(2000, 1, 1)
        out[t] = {
            "last_updated": last,
            "status":       row[2] if len(row) > 2 else "",
            "errors":       int(row[3]) if len(row) > 3 and row[3].isdigit() else 0,
        }
    return out


def save_status_batch(status_ws, updates: list[dict]):
    if not updates:
        return
    today    = date.today().isoformat()
    all_rows = status_ws.get_all_values()
    header   = all_rows[0] if all_rows else ["TICKER", "LAST_UPDATED", "STATUS", "CONSECUTIVE_ERRORS"]
    data     = {r[0].strip().upper(): list(r) for r in all_rows[1:] if r and r[0].strip()}

    for u in updates:
        ticker = u["ticker"].upper()
        row    = data.get(ticker, [ticker, "", "", ""])
        row    = list(row) + [""] * (4 - len(row))
        row[0], row[1], row[2], row[3] = ticker, today, u["status"], str(u["errors"])
        data[ticker] = row

    status_ws.clear()
    status_ws.update("A1", [header] + list(data.values()))
    log.info("Status saved (%d tickers tracked)", len(data))


# ── Alpha Vantage ─────────────────────────────────────────────────────────────
def av_get(function: str, symbol: str) -> Optional[dict]:
    """Returns dict on success, {} on empty/unknown, None on rate-limit."""
    params = {"function": function, "symbol": symbol, "apikey": AV_API_KEY}
    try:
        r = requests.get("https://www.alphavantage.co/query", params=params, timeout=20)
        r.raise_for_status()
        data = r.json()
        if "Note" in data or "Information" in data:
            log.warning("AV rate-limit signal for %s/%s", symbol, function)
            return None
        return data or {}
    except Exception as exc:
        log.error("AV error %s/%s: %s", symbol, function, exc)
        return {}


def pf(val) -> Optional[float]:
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def fmt(val: Optional[float]) -> str:
    """Always use dot as decimal separator, regardless of system locale."""
    if val is None:
        return ""
    return "{:.2f}".format(val)


# ── PER calculation ───────────────────────────────────────────────────────────
def price_at_quarter(monthly_ts: dict, fiscal_date: date) -> Optional[float]:
    """Closing price for the month of fiscal_date or next available month."""
    time_series = monthly_ts.get("Monthly Time Series", {})
    target_ym   = (fiscal_date.year, fiscal_date.month)
    candidates  = sorted(
        ((datetime.strptime(k, "%Y-%m-%d").date(), v) for k, v in time_series.items()),
        key=lambda x: x[0]
    )
    for d, ohlcv in candidates:
        if (d.year, d.month) >= target_ym:
            return pf(ohlcv.get("4. close"))
    return None


def calc_per(eps: Optional[float], price: Optional[float]) -> Optional[float]:
    if eps and eps > 0 and price and price > 0:
        return round(price / eps, 2)
    return None


def get_quarterly_pers(earnings: dict, monthly_ts: dict, n: int) -> list[Optional[float]]:
    """Compute PER for the last `n` quarters. Returns list of length n."""
    quarters = earnings.get("quarterlyEarnings", [])[:n]
    result   = []
    for q in quarters:
        eps = pf(q.get("reportedEPS"))
        try:
            fd = datetime.strptime(q["fiscalDateEnding"], "%Y-%m-%d").date()
        except (KeyError, ValueError):
            result.append(None)
            continue
        price = price_at_quarter(monthly_ts, fd)
        result.append(calc_per(eps, price))
    while len(result) < n:
        result.append(None)
    return result


def get_per_hist_3a(earnings: dict, monthly_ts: dict) -> Optional[float]:
    """Average PER over last 12 quarters (3 years)."""
    quarters = earnings.get("quarterlyEarnings", [])[:12]
    pers = []
    for q in quarters:
        eps = pf(q.get("reportedEPS"))
        try:
            fd = datetime.strptime(q["fiscalDateEnding"], "%Y-%m-%d").date()
        except (KeyError, ValueError):
            continue
        price = price_at_quarter(monthly_ts, fd)
        p = calc_per(eps, price)
        if p is not None:
            pers.append(p)
    return round(sum(pers) / len(pers), 2) if pers else None


# ── Build exactly 7 rows ──────────────────────────────────────────────────────
def build_rows(ticker: str, earnings: dict, monthly_ts: dict, overview: dict) -> list[list]:
    today  = date.today().strftime("%d/%m/%Y")
    pers   = get_quarterly_pers(earnings, monthly_ts, 5)
    hist3a = get_per_hist_3a(earnings, monthly_ts)
    peg    = pf(overview.get("PEGRatio"))

    rows = []
    for label, per in zip(["Q(n)", "Q(n-1)", "Q(n-2)", "Q(n-3)", "Q(n-4)"], pers):
        rows.append([today, ticker, "PER",         label,  fmt(per)])
    rows.append(    [today, ticker, "PER_HIST_3A", "Inst", fmt(hist3a)])
    rows.append(    [today, ticker, "PEG 5A",      "Q(n)", fmt(peg)])
    return rows   # exactly 7


# ── OUT sheet writer ──────────────────────────────────────────────────────────
def replace_ticker_rows(out_ws, ticker: str, new_rows: list[list]):
    """
    Replace the 7 rows for `ticker` in OUT sheet.
    Reads full sheet, filters out old rows for this ticker, appends new ones.
    """
    all_vals = out_ws.get_all_values()
    if not all_vals:
        kept = [OUT_HEADER]
    else:
        kept = [all_vals[0]]   # preserve header
        for row in all_vals[1:]:
            if len(row) >= 2 and row[1].strip().upper() == ticker.upper():
                continue       # drop old rows for this ticker
            kept.append(row)

    kept.extend(new_rows)
    out_ws.clear()
    out_ws.update("A1", kept)
    log.info("  → 7 rows written for %s in OUT", ticker)


# ── Scheduler ─────────────────────────────────────────────────────────────────
def choose_tickers(all_tickers: list[str], status: dict) -> list[str]:
    today = date.today()

    def sort_key(t):
        s      = status.get(t, {})
        age    = (today - s.get("last_updated", date(2000, 1, 1))).days
        errors = s.get("errors", 0)
        return (1000 if errors >= 3 else 0, -age)

    due = [
        t for t in sorted(all_tickers, key=sort_key)
        if (today - status.get(t, {}).get("last_updated", date(2000, 1, 1))).days >= UPDATE_CYCLE_DAYS
        or status.get(t, {}).get("errors", 0) > 0
    ]

    selected = due[:TICKERS_PER_DAY]
    log.info("Total tickers: %d | due: %d | processing today: %d",
             len(all_tickers), len(due), len(selected))
    return selected


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    log.info("=== AV Fetcher | sheet: %s | %d req/day | %d tickers/day ===",
             SHEET_ID, REQUESTS_PER_DAY, TICKERS_PER_DAY)

    wb        = open_workbook()
    ticker_ws = wb.worksheet(TICKER_SHEET)
    out_ws    = ensure_sheet(wb, OUT_SHEET)
    status_ws = ensure_sheet(wb, STATUS_SHEET, cols=4)

    # Ensure OUT has header
    ensure_out_header(out_ws)

    # Ensure STATUS has header
    if not status_ws.get_all_values():
        status_ws.update("A1:D1", [["TICKER", "LAST_UPDATED", "STATUS", "CONSECUTIVE_ERRORS"]])

    # Read tickers from col A (skip header row "TICKERS")
    all_tickers = [
        row[0].strip().upper()
        for row in ticker_ws.get_all_values()[1:]
        if row and row[0].strip()
    ]
    log.info("Tickers found in TICKERS tab: %d", len(all_tickers))

    status     = load_status(status_ws)
    to_process = choose_tickers(all_tickers, status)

    if not to_process:
        log.info("All tickers up-to-date. Nothing to do today.")
        return

    status_updates = []
    req_count      = 0
    rate_limit_hit = False

    for ticker in to_process:
        if rate_limit_hit:
            log.warning("Stopping — rate-limit hit.")
            break

        log.info("── %s ──", ticker)
        prev_errors = status.get(ticker, {}).get("errors", 0)

        # ── Call 1: EARNINGS ─────────────────────────────────────────────────
        earnings = av_get("EARNINGS", ticker)
        req_count += 1
        time.sleep(AV_DELAY_SECS)

        if earnings is None:
            rate_limit_hit = True
            status_updates.append({"ticker": ticker, "status": "RATE_LIMIT", "errors": prev_errors + 1})
            break
        if not earnings.get("quarterlyEarnings"):
            log.warning("  No quarterly earnings for %s — skipping", ticker)
            status_updates.append({"ticker": ticker, "status": "NO_DATA", "errors": prev_errors + 1})
            req_count += 2
            time.sleep(AV_DELAY_SECS * 2)
            continue

        # ── Call 2: TIME_SERIES_MONTHLY ───────────────────────────────────────
        monthly_ts = av_get("TIME_SERIES_MONTHLY", ticker)
        req_count += 1
        time.sleep(AV_DELAY_SECS)

        if monthly_ts is None:
            rate_limit_hit = True
            status_updates.append({"ticker": ticker, "status": "RATE_LIMIT", "errors": prev_errors + 1})
            break

        # ── Call 3: OVERVIEW ──────────────────────────────────────────────────
        overview = av_get("OVERVIEW", ticker)
        req_count += 1
        time.sleep(AV_DELAY_SECS)

        if overview is None:
            rate_limit_hit = True
            status_updates.append({"ticker": ticker, "status": "RATE_LIMIT", "errors": prev_errors + 1})
            break

        # ── Write 7 rows to OUT ───────────────────────────────────────────────
        rows = build_rows(ticker, earnings, monthly_ts, overview)
        replace_ticker_rows(out_ws, ticker, rows)
        status_updates.append({"ticker": ticker, "status": "OK", "errors": 0})
        log.info("  req used so far: %d/%d", req_count, REQUESTS_PER_DAY)

    save_status_batch(status_ws, status_updates)

    ok     = sum(1 for u in status_updates if u["status"] == "OK")
    errors = sum(1 for u in status_updates if u["status"] != "OK")
    log.info("=== Done | requests: %d | OK: %d | errors/skipped: %d | rate_limit: %s ===",
             req_count, ok, errors, rate_limit_hit)


if __name__ == "__main__":
    main()
