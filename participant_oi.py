#!/usr/bin/env python3
"""
NSE Participant-wise Open Interest -> Excel report.

Fetches the "F&O - Participant wise Open Interest" report from NSE for the 3 most
recent trading days and builds an Excel workbook mirroring the analyst layout:

  * Left block   "Futures & Options"           - day-over-day change per participant
                                                  (Added/Closed Longs & Shorts, Net Buy/Sell)
  * Middle block "Total Positions Carried"      - net position (Long-Short) for
                                                  TODAY / 1 DAY AGO / 2 DAYS AGO
  * Right block  "Positions Bought / Sold Today"- the net-change data pivoted by participant

Raw daily CSVs are cached under data/ so historical days are never re-fetched.
Run any time after ~7 PM IST; NSE publishes the file after market close.
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import sys
import urllib.parse
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from zoneinfo import ZoneInfo

    IST = ZoneInfo("Asia/Kolkata")
except Exception:  # pragma: no cover - fallback if tz data is unavailable
    from datetime import timezone

    IST = timezone(timedelta(hours=5, minutes=30))

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_FILE = ROOT / "participant_oi.xlsx"
HASH_FILE = ROOT / "report_data.hash"  # fingerprint of source data; drives commit-on-change

URL_TEMPLATE = "https://archives.nseindia.com/content/nsccl/fao_participant_oi_{ddmmyyyy}.csv"
INDEX_URL_TEMPLATE = "https://archives.nseindia.com/content/indices/ind_close_all_{ddmmyyyy}.csv"
NSE_HOME = "https://www.nseindia.com/"
NSE_OC_PAGE = "https://www.nseindia.com/option-chain"
OC_CONTRACT_URL = "https://www.nseindia.com/api/option-chain-contract-info?symbol={sym}"
OC_V3_URL = "https://www.nseindia.com/api/option-chain-v3?type=Indices&symbol={sym}&expiry={exp}"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
)
MAX_LOOKBACK_DAYS = 15  # safety bound when walking back over holidays
NUM_DAYS = 3            # today, 1 day ago, 2 days ago

# Participant rows: (CSV key, display label)
PARTICIPANTS = [
    ("Client", "Clients"),
    ("DII", "DIIs"),
    ("FII", "FIIs"),
    ("Pro", "Pro"),
]

# CSV column layout (0-based), header line 2 of the file.
COLS = {
    "Future Index Long": 1,
    "Future Index Short": 2,
    "Future Stock Long": 3,
    "Future Stock Short": 4,
    "Option Index Call Long": 5,
    "Option Index Put Long": 6,
    "Option Index Call Short": 7,
    "Option Index Put Short": 8,
    "Option Stock Call Long": 9,
    "Option Stock Put Long": 10,
    "Option Stock Call Short": 11,
    "Option Stock Put Short": 12,
}

# Instruments: (display name, long-col, short-col, long_is_bullish)
# long_is_bullish: a long future or long call is bullish, but a long PUT is bearish,
# so puts invert the sentiment colouring of every long/short/net cell.
INSTRUMENTS = [
    ("Index Futures", "Future Index Long", "Future Index Short", True),
    ("Index Call", "Option Index Call Long", "Option Index Call Short", True),
    ("Index Put", "Option Index Put Long", "Option Index Put Short", False),
    ("Stock Futures", "Future Stock Long", "Future Stock Short", True),
    ("Stock Calls", "Option Stock Call Long", "Option Stock Call Short", True),
    ("Stock Puts", "Option Stock Put Long", "Option Stock Put Short", False),
]

# --------------------------------------------------------------------------- #
# Fetch + parse
# --------------------------------------------------------------------------- #
def _fetch_raw(d: date) -> str | None:
    """Return the CSV text for date d, or None if NSE has no file (404 = holiday)."""
    cache = DATA_DIR / f"fao_participant_oi_{d.isoformat()}.csv"
    if cache.exists():
        return cache.read_text()

    url = URL_TEMPLATE.format(ddmmyyyy=d.strftime("%d%m%Y"))
    try:
        resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=30)
    except requests.RequestException as exc:
        print(f"  ! network error for {d}: {exc}", file=sys.stderr)
        return None

    if resp.status_code == 404:
        return None
    resp.raise_for_status()

    text = resp.text
    if "Participant wise Open Interest" not in text:
        return None

    DATA_DIR.mkdir(exist_ok=True)
    cache.write_text(text)
    return text


def _parse(text: str) -> dict[str, dict[str, int]]:
    """Parse CSV text into {participant: {column-name: value}}."""
    reader = csv.reader(io.StringIO(text))
    rows = [r for r in reader if r]
    out: dict[str, dict[str, int]] = {}
    for row in rows:
        key = row[0].strip()
        if key not in ("Client", "DII", "FII", "Pro", "TOTAL"):
            continue
        out[key] = {name: int(row[idx].strip() or 0) for name, idx in COLS.items()}
    return out


def load_recent_days(n: int = NUM_DAYS) -> list[tuple[date, dict]]:
    """Return the n most recent trading days as (date, parsed) newest first."""
    days: list[tuple[date, dict]] = []
    d = datetime.now(IST).date()
    steps = 0
    while len(days) < n and steps < MAX_LOOKBACK_DAYS:
        text = _fetch_raw(d)
        if text:
            parsed = _parse(text)
            if parsed:
                print(f"  loaded {d}")
                days.append((d, parsed))
        d -= timedelta(days=1)
        steps += 1
    if len(days) < n:
        raise SystemExit(
            f"Only found {len(days)} trading day(s) in the last {MAX_LOOKBACK_DAYS} days; "
            f"need {n}. NSE may not have published today's file yet."
        )
    return days


def net(day: dict, participant: str, long_col: str, short_col: str) -> int:
    """Net position = Long - Short for a participant/instrument on a given day."""
    p = day[participant]
    return p[long_col] - p[short_col]


def fetch_nifty_ohlc(d: date, max_back: int = 7):
    """Return (date, open, high, low, close) for Nifty 50 on/just before d, or None."""
    cur = d
    for _ in range(max_back):
        cache = DATA_DIR / f"ind_close_all_{cur.isoformat()}.csv"
        text = cache.read_text() if cache.exists() else None
        if text is None:
            url = INDEX_URL_TEMPLATE.format(ddmmyyyy=cur.strftime("%d%m%Y"))
            try:
                resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=30)
            except requests.RequestException:
                resp = None
            if resp is not None and resp.status_code == 200 and "Index Name" in resp.text:
                text = resp.text
                DATA_DIR.mkdir(exist_ok=True)
                cache.write_text(text)
        if text:
            for row in csv.reader(io.StringIO(text)):
                if row and row[0].strip() == "Nifty 50":
                    return (cur, float(row[2]), float(row[3]), float(row[4]), float(row[5]))
        cur -= timedelta(days=1)
    return None


def pivot_levels(high: float, low: float, close: float) -> dict[str, float]:
    """Classic floor-trader pivot points + Central Pivot Range for the next session."""
    pp = (high + low + close) / 3
    return {
        "pp": pp,
        "r1": 2 * pp - low,
        "s1": 2 * pp - high,
        "r2": pp + (high - low),
        "s2": pp - (high - low),
        "r3": high + 2 * (pp - low),
        "s3": low - 2 * (high - pp),
        "bc": (high + low) / 2,
        "tc": 2 * pp - (high + low) / 2,
    }


def fetch_nifty_option_chain(symbol: str = "NIFTY") -> dict | None:
    """Live option chain for the nearest (this-week) NIFTY expiry, or None on failure.

    Returns {expiry, spot, pcr, max_pain, strikes:[{strike, ce_oi, ce_chg, ce_vol,
    pe_oi, pe_chg, pe_vol}]}. OI/volume are as published by NSE (end-of-day after close).
    """
    try:
        s = requests.Session()
        s.headers.update({
            "User-Agent": USER_AGENT,
            "Accept-Language": "en-US,en;q=0.9",
            "Sec-Fetch-Dest": "document", "Sec-Fetch-Mode": "navigate", "Sec-Fetch-Site": "none",
        })
        s.get(NSE_HOME, timeout=30)        # seed Akamai cookies
        s.get(NSE_OC_PAGE, timeout=30)
        api_headers = {
            "Accept": "application/json, text/plain, */*",
            "Referer": NSE_OC_PAGE,
            "Sec-Fetch-Dest": "empty", "Sec-Fetch-Mode": "cors", "Sec-Fetch-Site": "same-origin",
        }
        info = s.get(OC_CONTRACT_URL.format(sym=symbol), headers=api_headers, timeout=30).json()
        expiry = info["expiryDates"][0]  # nearest upcoming expiry = this week
        url = OC_V3_URL.format(sym=symbol, exp=urllib.parse.quote(expiry))
        data = s.get(url, headers=api_headers, timeout=30).json()
    except Exception as exc:  # network / bot-block / shape change -> skip gracefully
        print(f"  ! option chain unavailable: {type(exc).__name__}: {exc}", file=sys.stderr)
        return None

    rec = data.get("records", {})
    spot = rec.get("underlyingValue")
    strikes = []
    for q in rec.get("data", []):
        ce, pe = q.get("CE"), q.get("PE")
        strikes.append({
            "strike": q["strikePrice"],
            "ce_oi": ce["openInterest"] if ce else 0,
            "ce_chg": ce["changeinOpenInterest"] if ce else 0,
            "ce_vol": ce["totalTradedVolume"] if ce else 0,
            "pe_oi": pe["openInterest"] if pe else 0,
            "pe_chg": pe["changeinOpenInterest"] if pe else 0,
            "pe_vol": pe["totalTradedVolume"] if pe else 0,
        })
    if not strikes or spot is None:
        return None

    tot_ce = sum(r["ce_oi"] for r in strikes)
    tot_pe = sum(r["pe_oi"] for r in strikes)
    pcr = (tot_pe / tot_ce) if tot_ce else 0.0

    # Max pain = expiry strike minimising total payout by option writers.
    def payout(k):
        return (sum(r["ce_oi"] * max(0, k - r["strike"]) for r in strikes)
                + sum(r["pe_oi"] * max(0, r["strike"] - k) for r in strikes))
    max_pain = min((r["strike"] for r in strikes), key=payout)

    return {"expiry": expiry, "spot": spot, "pcr": pcr, "max_pain": max_pain, "strikes": strikes}


def _strength(vol: int, mx: int) -> str:
    """Grade a level's conviction from its traded volume relative to the strongest shown."""
    f = (vol / mx) if mx else 0.0
    return "High" if f >= 0.66 else "Medium" if f >= 0.33 else "Low"


def analyze_oi_sr(oc: dict, n: int = 3) -> dict:
    """Top-n Call-OI resistance walls and Put-OI support walls, graded by volume."""
    spot, strikes = oc["spot"], oc["strikes"]
    atm = min(strikes, key=lambda r: abs(r["strike"] - spot))["strike"]
    res = sorted([r for r in strikes if r["strike"] >= atm and r["ce_oi"] > 0],
                 key=lambda r: r["ce_oi"], reverse=True)[:n]
    sup = sorted([r for r in strikes if r["strike"] <= atm and r["pe_oi"] > 0],
                 key=lambda r: r["pe_oi"], reverse=True)[:n]
    res.sort(key=lambda r: r["strike"])              # nearest resistance above spot first
    sup.sort(key=lambda r: r["strike"], reverse=True)  # nearest support below spot first
    mx = max([r["ce_vol"] for r in res] + [r["pe_vol"] for r in sup] + [1])
    return {"atm": atm, "resistance": res, "support": sup, "max_vol": mx}


# --------------------------------------------------------------------------- #
# Excel styling
# --------------------------------------------------------------------------- #
NAVY = PatternFill("solid", fgColor="1F2A44")
SUBHEAD = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
INSTR_FILL = PatternFill("solid", fgColor="EDEDED")

WHITE_BOLD = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)
GREEN = Font(color="0B7A34")
RED = Font(color="C00000")
GREEN_B = Font(color="0B7A34", bold=True)
RED_B = Font(color="C00000", bold=True)
GREY = Font(color="808080", italic=True)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

IND_FMT = "#,##,##0;-#,##,##0;0"  # Indian digit grouping, kept numeric
NIFTY_FMT = "#,##,##0.00"          # index level with two decimals


def _put(ws, r, c, value, *, font=None, fill=None, align=None, fmt=None, border=True):
    cell = ws.cell(row=r, column=c, value=value)
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align
    if fmt:
        cell.number_format = fmt
    if border:
        cell.border = BORDER
    return cell


def _colour(bullish):
    return GREEN if bullish else RED


def _long_label(delta, long_bullish=True):
    # Adding to a bullish long is bullish; for puts (long_bullish=False) it inverts.
    label = "Added Longs" if delta >= 0 else "Closed Longs"
    bullish = (delta >= 0) if long_bullish else (delta < 0)
    return label, _colour(bullish)


def _short_label(delta, long_bullish=True):
    # A short is the opposite stance of the long, so its sentiment is inverted again.
    label = "Added Shorts" if delta >= 0 else "Closed Shorts"
    bullish = (delta < 0) if long_bullish else (delta >= 0)
    return label, _colour(bullish)


def _net_label(delta, long_bullish=True):
    # Net buying is bullish for futures/calls, bearish for puts (buying puts = bearish).
    label = "Bought Net" if delta >= 0 else "Sold Net"
    bullish = (delta >= 0) if long_bullish else (delta < 0)
    return label, _colour(bullish)


# --------------------------------------------------------------------------- #
# Build workbook
# --------------------------------------------------------------------------- #
def build(days: list[tuple[date, dict]], ohlc, oc: dict | None, path: Path) -> None:
    today, d1, d2 = days[0], days[1], days[2]
    (dt_today, day_today) = today
    (dt_1, day_1) = d1
    (dt_2, day_2) = d2

    wb = Workbook()
    ws = wb.active
    ws.title = "Participant OI"
    ws.sheet_view.showGridLines = False

    # ---- title row -------------------------------------------------------- #
    _put(ws, 1, 1, "Futures & Options", font=WHITE_BOLD, fill=NAVY, align=LEFT)
    for c in range(2, 8):
        _put(ws, 1, c, None, fill=NAVY)
    _put(ws, 1, 9, "Total Positions Carried", font=WHITE_BOLD, fill=NAVY, align=CENTER)
    for c in (10, 11, 12):
        _put(ws, 1, c, None, fill=NAVY)
    _put(ws, 1, 14, "Positions Bought / Sold Today", font=WHITE_BOLD, fill=NAVY, align=CENTER)
    for c in (15, 16, 17):
        _put(ws, 1, c, None, fill=NAVY)

    r = 2
    # Each instrument section occupies: 1 sub-header + 4 participants + 1 total = 6 rows
    for instr_name, long_col, short_col, long_bullish in INSTRUMENTS:
        # ---- LEFT block sub-headers ---- #
        _put(ws, r, 1, "", fill=INSTR_FILL)
        _put(ws, r, 2, f"{instr_name} Longs", font=BOLD, fill=SUBHEAD, align=CENTER)
        _put(ws, r, 3, "", fill=SUBHEAD)
        _put(ws, r, 4, f"{instr_name} Shorts", font=BOLD, fill=SUBHEAD, align=CENTER)
        _put(ws, r, 5, "", fill=SUBHEAD)
        _put(ws, r, 6, "Net Buy / Sell for Today", font=BOLD, fill=SUBHEAD, align=CENTER)
        _put(ws, r, 7, "", fill=SUBHEAD)
        # ---- MIDDLE block sub-headers ---- #
        _put(ws, r, 9, instr_name, font=BOLD, fill=SUBHEAD, align=LEFT)
        _put(ws, r, 10, "TODAY", font=BOLD, fill=SUBHEAD, align=CENTER)
        _put(ws, r, 11, "1 DAY AGO", font=BOLD, fill=SUBHEAD, align=CENTER)
        _put(ws, r, 12, "2 DAYS AGO", font=BOLD, fill=SUBHEAD, align=CENTER)
        r += 1

        tot_dlong = tot_dshort = 0
        for csv_key, disp in PARTICIPANTS:
            long_t = day_today[csv_key][long_col]
            short_t = day_today[csv_key][short_col]
            long_y = day_1[csv_key][long_col]
            short_y = day_1[csv_key][short_col]
            d_long = long_t - long_y
            d_short = short_t - short_y
            d_net = d_long - d_short
            tot_dlong += d_long
            tot_dshort += d_short

            l_lbl, l_font = _long_label(d_long, long_bullish)
            s_lbl, s_font = _short_label(d_short, long_bullish)
            n_lbl, n_font = _net_label(d_net, long_bullish)

            _put(ws, r, 1, disp, font=BOLD, align=LEFT)
            _put(ws, r, 2, l_lbl, font=l_font, align=LEFT)
            _put(ws, r, 3, d_long, font=l_font, align=RIGHT, fmt=IND_FMT)
            _put(ws, r, 4, s_lbl, font=s_font, align=LEFT)
            _put(ws, r, 5, d_short, font=s_font, align=RIGHT, fmt=IND_FMT)
            _put(ws, r, 6, n_lbl, font=n_font, align=LEFT)
            _put(ws, r, 7, d_net, font=n_font, align=RIGHT, fmt=IND_FMT)

            # ---- MIDDLE block: net positions carried ---- #
            _put(ws, r, 9, disp, font=BOLD, align=LEFT)
            _put(ws, r, 10, net(day_today, csv_key, long_col, short_col), align=RIGHT, fmt=IND_FMT)
            _put(ws, r, 11, net(day_1, csv_key, long_col, short_col), align=RIGHT, fmt=IND_FMT)
            _put(ws, r, 12, net(day_2, csv_key, long_col, short_col), align=RIGHT, fmt=IND_FMT)
            r += 1

        # ---- Total row ---- #
        tot_dnet = tot_dlong - tot_dshort
        _put(ws, r, 1, "Total", font=BOLD, fill=TOTAL_FILL, align=LEFT)
        _put(ws, r, 2, "", fill=TOTAL_FILL)
        _put(ws, r, 3, tot_dlong, font=BOLD, fill=TOTAL_FILL, align=RIGHT, fmt=IND_FMT)
        _put(ws, r, 4, "", fill=TOTAL_FILL)
        _put(ws, r, 5, tot_dshort, font=BOLD, fill=TOTAL_FILL, align=RIGHT, fmt=IND_FMT)
        _put(ws, r, 6, "", fill=TOTAL_FILL)
        _put(ws, r, 7, tot_dnet, font=BOLD, fill=TOTAL_FILL, align=RIGHT, fmt=IND_FMT)
        for c in (9, 10, 11, 12):
            _put(ws, r, c, "Total" if c == 9 else "", font=BOLD, fill=TOTAL_FILL, align=LEFT)
        r += 1

    # ---- RIGHT block: Positions Bought / Sold Today, grouped by participant --- #
    rr = 2
    for csv_key, disp in PARTICIPANTS:
        for instr_name, long_col, short_col, long_bullish in INSTRUMENTS:
            d_net = (
                (day_today[csv_key][long_col] - day_1[csv_key][long_col])
                - (day_today[csv_key][short_col] - day_1[csv_key][short_col])
            )
            n_lbl, n_font = _net_label(d_net, long_bullish)
            _put(ws, rr, 14, disp, font=BOLD, align=LEFT)
            _put(ws, rr, 15, n_lbl, font=n_font, align=LEFT)
            _put(ws, rr, 16, instr_name, align=LEFT)
            _put(ws, rr, 17, d_net, font=n_font, align=RIGHT, fmt=IND_FMT)
            rr += 1

    # ---- Nifty support & resistance (bottom-right, below the right block) ---- #
    sr_last = rr
    if ohlc:
        ndate, o, h, l, c = ohlc
        p = pivot_levels(h, l, c)
        s = rr + 2  # leave a blank row under the right block

        _put(ws, s, 14, "Nifty 50  -  Support & Resistance", font=WHITE_BOLD, fill=NAVY, align=CENTER)
        for cc in (15, 16, 17):
            _put(ws, s, cc, None, fill=NAVY)
        _put(ws, s + 1, 14, f"Next session  -  based on {ndate:%d-%b-%Y} OHLC",
             font=BOLD, fill=SUBHEAD, align=CENTER)
        for cc in (15, 16, 17):
            _put(ws, s + 1, cc, "", fill=SUBHEAD)

        # OHLC line
        _put(ws, s + 2, 14, "Open", font=BOLD, align=LEFT)
        _put(ws, s + 2, 15, o, align=RIGHT, fmt=NIFTY_FMT)
        _put(ws, s + 2, 16, "High", font=BOLD, align=LEFT)
        _put(ws, s + 2, 17, h, align=RIGHT, fmt=NIFTY_FMT)
        _put(ws, s + 3, 14, "Low", font=BOLD, align=LEFT)
        _put(ws, s + 3, 15, l, align=RIGHT, fmt=NIFTY_FMT)
        _put(ws, s + 3, 16, "Close", font=BOLD, align=LEFT)
        _put(ws, s + 3, 17, c, align=RIGHT, fmt=NIFTY_FMT)

        # Resistances / pivot / supports (col 14-15) beside CPR (col 16-17)
        levels = [
            ("R3", p["r3"], RED_B), ("R2", p["r2"], RED_B), ("R1", p["r1"], RED_B),
            ("Pivot", p["pp"], BOLD),
            ("S1", p["s1"], GREEN_B), ("S2", p["s2"], GREEN_B), ("S3", p["s3"], GREEN_B),
        ]
        cpr = [("CPR Top", p["tc"], BOLD), ("Pivot", p["pp"], BOLD), ("CPR Bottom", p["bc"], BOLD)]
        for i, (name, val, font) in enumerate(levels):
            row = s + 4 + i
            _put(ws, row, 14, name, font=font, align=LEFT)
            _put(ws, row, 15, val, font=font, align=RIGHT, fmt=NIFTY_FMT)
            if i < len(cpr):
                cname, cval, cfont = cpr[i]
                _put(ws, row, 16, cname, font=cfont, align=LEFT)
                _put(ws, row, 17, cval, font=cfont, align=RIGHT, fmt=NIFTY_FMT)
            else:
                _put(ws, row, 16, "", border=False)
                _put(ws, row, 17, "", border=False)
        sr_last = s + 4 + len(levels)

    # ---- Nifty option-chain OI support & resistance, graded by volume ---- #
    if oc:
        a = analyze_oi_sr(oc)
        mx = a["max_vol"]
        o = sr_last + 2

        def _merge(row, text, font, fill):
            ws.merge_cells(start_row=row, start_column=14, end_row=row, end_column=18)
            _put(ws, row, 14, text, font=font, fill=fill, align=CENTER)
            for cc in range(15, 19):
                _put(ws, row, cc, None, fill=fill)

        _merge(o, f"Nifty Option-Chain S/R   -   Weekly expiry {oc['expiry']}", WHITE_BOLD, NAVY)
        _merge(o + 1,
               f"Spot {oc['spot']:,.0f}    PCR(OI) {oc['pcr']:.2f}    Max Pain {oc['max_pain']:,.0f}",
               BOLD, SUBHEAD)

        def _headers(row, oi_label):
            for cc, txt in zip(range(14, 19), ["Strike", oi_label, "Chg OI", "Volume", "Strength"]):
                _put(ws, row, cc, txt, font=BOLD, fill=INSTR_FILL, align=CENTER)

        def _level_row(row, rec, oi, chg, vol, strike_font):
            lbl = _strength(vol, mx)
            sfont = BOLD if lbl == "High" else (GREY if lbl == "Low" else None)
            _put(ws, row, 14, rec["strike"], font=strike_font, align=RIGHT, fmt="#,##0")
            _put(ws, row, 15, oi, align=RIGHT, fmt=IND_FMT)
            _put(ws, row, 16, chg, font=(GREEN if chg >= 0 else RED), align=RIGHT, fmt=IND_FMT)
            _put(ws, row, 17, vol, align=RIGHT, fmt=IND_FMT)
            _put(ws, row, 18, lbl, font=sfont, align=CENTER)

        row = o + 2
        _merge(row, "RESISTANCE   -   Call OI walls (higher = stronger cap)", RED_B, SUBHEAD)
        _headers(row + 1, "Call OI")
        row += 2
        for rec in a["resistance"]:
            _level_row(row, rec, rec["ce_oi"], rec["ce_chg"], rec["ce_vol"], RED_B)
            row += 1

        _merge(row, "SUPPORT   -   Put OI walls (higher = stronger floor)", GREEN_B, SUBHEAD)
        _headers(row + 1, "Put OI")
        row += 2
        for rec in a["support"]:
            _level_row(row, rec, rec["pe_oi"], rec["pe_chg"], rec["pe_vol"], GREEN_B)
            row += 1
        sr_last = row

    # ---- footer with the dates used ---- #
    foot = max(r, rr, sr_last) + 1
    _put(
        ws, foot, 1,
        f"TODAY = {dt_today:%d-%b-%Y}    1 DAY AGO = {dt_1:%d-%b-%Y}    "
        f"2 DAYS AGO = {dt_2:%d-%b-%Y}    (generated {datetime.now(IST):%d-%b-%Y %H:%M IST})",
        font=Font(italic=True, color="808080"), border=False,
    )

    # ---- column widths ---- #
    widths = {
        1: 9, 2: 13, 3: 11, 4: 13, 5: 11, 6: 16, 7: 11, 8: 2,
        9: 15, 10: 12, 11: 12, 12: 12, 13: 2,
        14: 10, 15: 12, 16: 15, 17: 12, 18: 10,
    }
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"

    wb.save(path)


def data_fingerprint(days: list[tuple[date, dict]], ohlc) -> str:
    """Deterministic hash of the report's source data (ignores generation time).

    Covers the 3 days of participant OI and the Nifty OHLC - the archive-based,
    date-stamped inputs. Unchanged inputs -> identical hash -> no new commit.
    """
    payload = {
        "days": [[d.isoformat(), day] for d, day in days],
        "ohlc": [ohlc[0].isoformat(), *ohlc[1:]] if ohlc else None,
    }
    blob = json.dumps(payload, sort_keys=True, default=str)
    return hashlib.sha256(blob.encode()).hexdigest()


def main() -> None:
    today = datetime.now(IST).date()

    # Skip weekends & market holidays: NSE publishes no participant file that day.
    if _fetch_raw(today) is None:
        print(f"{today:%d-%b-%Y}: weekend/market holiday - no NSE data published, skipping run.")
        return

    print("Fetching NSE participant-wise OI...")
    days = load_recent_days()
    ohlc = fetch_nifty_ohlc(days[0][0])
    oc = fetch_nifty_option_chain("NIFTY")
    print(f"Building {OUTPUT_FILE.name}...")
    build(days, ohlc, oc, OUTPUT_FILE)
    HASH_FILE.write_text(data_fingerprint(days, ohlc) + "\n")
    print(f"Done -> {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
