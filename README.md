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
- **India VIX** - closing VIX for the same 5 sessions (middle block, below participant tables).
- **Nifty SRT** - Speculation Ratio Territory (Nifty close ÷ 124-day SMA) for the same
  5 sessions, directly below VIX; green ≤0.9 (accumulation band), red ≥1.3 (exit band).
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

The Python environment (python3 + `openpyxl` + `requests`) is declared
reproducibly in `flake.nix` - no global/pip installs required.

```bash
nix run          # generate the report
nix develop      # shell with the pinned python for hacking
python3 participant_oi.py   # inside the dev shell
```

`run_nightly.sh` builds this env via `nix build .#pythonEnv --out-link .nix-python`
(a GC root, so it survives `nix-collect-garbage`) and runs the script with it.

## Scheduling (macOS, 8:30 PM + 10:30 PM IST on trading days)

A LaunchAgent (`com.marketanalysis.participantoi`) fires `run_nightly.sh` **twice
per weekday (Mon-Fri)**: a primary run at 20:30 local time (8:30 PM IST) plus a
fallback at 22:30 (10:30 PM IST). NSE usually publishes the participant file
6:30-7:30 PM IST but occasionally after 8:30 PM (seen as late as ~8:57 PM), so the
8:30 run can miss those days; the 10:30 run then catches them. Because commits are
gated on `report_data.hash` (see below), the second fire is a harmless no-op
whenever the first already captured the day's data. The wrapper:

1. Runs `participant_oi.py`, which **no-ops on market holidays** (if NSE has
   published no participant file for the day, nothing is regenerated).
2. Commits and pushes the workbook **only when the source data actually changed** -
   gated on `report_data.hash`, a fingerprint of the data (not the generation
   timestamp), so re-runs and non-trading days produce no noise commits.

Weekends are excluded at the launchd level (weekday-only triggers); weekday
holidays are handled by the runtime check above.

Install: copy `com.marketanalysis.participantoi.plist` to `~/Library/LaunchAgents/`
and load with `launchctl bootstrap gui/$(id -u) <plist>`.
