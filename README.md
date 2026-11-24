# Rajasthan Panchayat Election Data Scraper & Explorer

An end-to-end data pipeline that collects **Contesting Sarpanch** candidate records for every Panchayat Samiti in Rajasthan from the official [Rajasthan State Election Commission](https://sec.rajasthan.gov.in/GramPanchayatDetails.aspx) portal, merges them into one structured Excel workbook, and serves them through a one-page web explorer.

**Coverage:** 41 Districts · 457 Panchayat Samitis · 14,370 Gram Panchayats · 82,824 candidates (General Election, OCT–NOV 2026)

## Highlights

- **Browser automation for an ASP.NET WebForms site** (Playwright): cascading District → Panchayat Samiti → Gram Panchayat postbacks, ViewState handling, UpdatePanel synchronization.
- **Automatic discovery** of all districts and Panchayat Samitis instead of hard-coded lists.
- **Reliability:** retries with exponential backoff, timeouts, page resets, crash-safe atomic checkpoints, and resume after interruption.
- **Data accuracy:** strict 17-column parsing by cell index, 100% cell-for-cell parity check against the portal's own "Export to Excel", no deduplication or alteration of source values; anomalies are reported, not hidden.
- **Polite scraping:** single sequential session with 1.5–2.5 s delays.
- **Reporting:** formatted Excel workbook (filters, frozen panes, summary sheets), progress tracker, and error / missing-record report.
- **Web explorer:** filter by District, Panchayat Samiti, Gram Panchayat or name; download any selection as Excel/CSV; re-fetch a single Panchayat Samiti live from the portal.

## Project Structure

| File | Purpose |
|---|---|
| `main.py` | CLI: `--discover`, `--smoke-test N`, `--run`, `--build-excel`, `--all` |
| `scraper.py` | Playwright scraping engine (retries, rate limiting, parity check) |
| `excel_builder.py` | Combined Excel workbook + error/missing and summary reports |
| `progress_tracker.py` | Atomic JSON/CSV progress tracker for resumable runs |
| `schema.py` | Column schema and lossless value parsers |
| `server.py` | One-page web explorer (Python standard-library HTTP server) |
| `static/` | Front end (`index.html`, `styles.css`, `app.js`) |
| `config.json` | URL, election filters, skip list, timeouts, delays, output paths |

## Setup

```bash
pip install -r requirements.txt
playwright install chromium   # or have Google Chrome installed (used automatically as fallback)
```

## Usage

```bash
python3 main.py --discover                 # find all districts & Panchayat Samitis
python3 main.py --smoke-test 4 --build-excel   # quick test on 4 Panchayat Samitis
python3 main.py --run --build-excel        # full / resumed scrape + final workbook
python3 main.py --build-excel              # rebuild workbook & reports only

python3 server.py                          # web explorer at http://127.0.0.1:8000
```

## Excel Layout

The `Contesting Candidates` sheet starts with **S. No. (Overall)** (1 … 82,824), **District** and **Panchayat Samiti**, followed by the original 17 columns. The portal's own **Sr. No.** restarts at 1 for every Panchayat Samiti, so it is kept as *Sr. No. (within Panchayat Samiti)*. Use the header filter buttons to pick any District, Panchayat Samiti or Gram Panchayat.

## Data & Privacy

The scraped data contains candidates' personal details (mobile numbers, e-mails, asset declarations). It is excluded from version control via `.gitignore` (`output/`, `*.xlsx`, `*.csv`). Share the **code**, not the data, in public repositories, and follow the source website's terms of use.
