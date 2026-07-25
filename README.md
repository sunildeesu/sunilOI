# Market Analysis - Participant-wise OI

Builds a daily Excel report from NSE F&O data, mirroring the analyst layout for
**participant-wise open interest** plus two independent **Nifty support/resistance** views.

## What it produces

A single workbook, `participant_oi.xlsx`, overwritten each run with the 3 most
recent trading days:

- **Futures & Options** - per participant (Clients / DIIs / FIIs / Pro) and
  instrument (Index & Stock Futures/Calls/Puts): day-over-day change labelled
  *Added/Closed Longs & Shorts* and *Bought/Sold Net*, colour-coded by market
  sentiment (bullish = green, bearish = red).
- **Total Positions Carried** - net position (Long − Short) for TODAY / 1 DAY AGO / 2 DAYS AGO.
- **Positions Bought / Sold Today** - the net-change data pivoted by participant.
- **Nifty S/R (pivot points)** - classic floor-trader pivots (R1–R3, S1–S3) + CPR
  from the latest session's OHLC.
- **Nifty S/R (option-chain OI walls)** - top Call-OI resistance and Put-OI support
  strikes for the current weekly expiry, graded **High/Medium/Low by traded volume**,
  with PCR and Max Pain.

## Data sources (NSE)

| Data | Endpoint |
|------|----------|
| Participant-wise OI | `archives.nseindia.com/content/nsccl/fao_participant_oi_DDMMYYYY.csv` |
| Nifty OHLC | `archives.nseindia.com/content/indices/ind_close_all_DDMMYYYY.csv` |
| Option chain (weekly) | `nseindia.com/api/option-chain-v3` (live, cookie session) |

Weekends/holidays return 404, so the script walks back to the 3 most recent
trading days automatically. Raw daily files are cached under `data/`.

## Usage

```bash
pip3 install openpyxl requests   # pandas optional
python3 participant_oi.py
```

## Scheduling (macOS, 10:30 PM IST nightly)

A LaunchAgent (`com.marketanalysis.participantoi`) runs the script daily at 22:30
local time. See `com.marketanalysis.participantoi.plist` in this repo for the
template; install it to `~/Library/LaunchAgents/` and load with
`launchctl bootstrap gui/$(id -u) <plist>`.
