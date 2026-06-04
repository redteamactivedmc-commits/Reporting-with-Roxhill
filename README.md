# Roxhill Coverage Agent

### Active DMC — Automated Media Monitoring Pipeline

Reads Roxhill alert emails from Outlook and feeds them directly into
the ADMC media monitor report system. No more manual URL entry.

---

## Architecture

```
Outlook Inbox
      │
      │  monitoring-emails@roxhillmedia.com
      ▼
┌─────────────────────┐
│   outlook_fetcher   │  Calls Microsoft Graph API
│                     │  Filters by sender + date window
│                     │  Returns raw email HTML bodies
└────────┬────────────┘
         │
         ▼
┌─────────────────────┐
│   roxhill_parser    │  Parses Roxhill HTML structure
│                     │  Decodes tracking URLs → real article URLs
│                     │  Extracts: headline, publication, tier,
│                     │  date, reach, Online/Print, language
└────────┬────────────┘
         │
         ▼
┌─────────────────────┐
│   roxhill_agent     │  Orchestrates everything
│   (main)            │  Groups articles by client
│                     │  Applies tier and client filters
│                     │  Exports JSON
└────────┬────────────┘
         │
         ├──▶ create_coverage_24h()  →  ADMC_<Client>_<Campaign>_<Date>.docx
         ├──▶ create_snapshot()      →  ADMC_<Client>_<Pub>_<Date>.docx (per article)
         └──▶ update_tracker()       →  ADMC_clients_<Client>tracker.xlsx
```

---

## File structure

```
Reporting-with-Roxhill/
├── roxhill_agent.py         ← Main orchestrator (run this)
├── roxhill_parser.py        ← Roxhill HTML parser + URL decoder
├── outlook_fetcher.py       ← Microsoft Graph API email fetcher
├── media_monitor/           ← Report generation package
│   ├── __init__.py
│   ├── __main__.py          ← CLI entry point
│   ├── config.py            ← Central configuration
│   ├── snapshot.py          ← Command 1: article snapshots
│   ├── coverage_24h.py      ← Command 2: 24h coverage reports
│   ├── tracker.py           ← Command 3: Excel tracker
│   └── utils/
│       ├── browser.py       ← Headless Chrome screenshots
│       ├── excel_helper.py  ← Excel workbook utilities
│       ├── file_naming.py   ← ADMC naming conventions
│       └── logo_handler.py  ← Logo insertion for Word docs
├── assets/
│   ├── active_logo.png      ← ADMC/Active logo
│   └── client_logos/        ← Client logos (e.g. knowbe4.png)
├── requirements.txt
├── requirements_roxhill.txt
├── .env.example             ← Credentials template
├── coverage_items.json      ← Example data for 24h command
└── tracker_entries.json     ← Example data for tracker command
```

---

## Quick start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Set up credentials
cp .env.example .env
# Edit .env with your Azure credentials

# 3. Test without creating files
python roxhill_agent.py --dry-run

# 4. Run for a specific client, last 24h
python roxhill_agent.py --client uipath

# 5. Full run, all clients, all report types
python roxhill_agent.py

# 6. Tracker update only, Tier A articles only
python roxhill_agent.py --output tracker --tier A
```

---

## Media Monitor CLI (standalone)

```bash
# Snapshot
python -m media_monitor snapshot \
  --url "https://example.com/article" \
  --client "KnowBe4" \
  --magazine "SecurityMEA"

# 24h Coverage Report
python -m media_monitor coverage-24h \
  --client "KnowBe4" \
  --campaign "Phish Alert Button Launch" \
  --items-file coverage_items.json

# Excel Tracker
python -m media_monitor tracker update \
  --client "KnowBe4" \
  --entries-file tracker_entries.json
```

---

## CLI options (roxhill_agent.py)

| Flag | Values | Default | Description |
|------|--------|---------|-------------|
| `--client` | any alias key | all | Filter by client |
| `--hours` | integer | 24 | Lookback window |
| `--output` | 24h / snapshot / tracker / all | all | Which reports to generate |
| `--tier` | any / A / AB | any | Minimum tier to include |
| `--dry-run` | flag | off | Parse and print, no files |
| `--no-json` | flag | off | Skip JSON export |

---

## Client alias keys

| Alias key | Maps to |
|-----------|---------|
| uipath | UiPath |
| illumio | Illumio |
| denodo | Denodo |
| phosphorus | Phosphorus Cybersecurity |
| knowbe4 | KnowBe4 |
| netscout | NETSCOUT |
| dynatrace | Dynatrace |
| qlik | Qlik |
| commscope | CommScope |
| heidrick | Heidrick & Struggles |
| levelinfinite | Level Infinite |
| emerson | Emerson |
| jpmorgan | JP Morgan Private Bank |
| ciena | Ciena |
| cequence | Cequence |

To add a new client, edit the `CLIENT_ALIAS_MAP` dict in `roxhill_parser.py`.

---

## Output file naming

| Command | Output File |
|---------|-------------|
| snapshot | `ADMC_<Client>_<Magazine>_<DDMMYYYY>.docx` |
| coverage-24h | `ADMC_<Client>_<Campaign>_<DDMMYYYY>.docx` |
| tracker update | `ADMC_clients_<Client>tracker.xlsx` |

---

## Setup: logos

```
assets/
├── active_logo.png              ← Drop your Active/ADMC logo here
└── client_logos/
    ├── knowbe4.png              ← Named as client name (lowercase)
    ├── phosphorus.png
    └── denodo.png
```

Both logos auto-appear at the top of every generated document.

---

## One-time Azure setup

1. Go to portal.azure.com → Azure Active Directory → App registrations
2. New registration → name: "ADMC Roxhill Agent"
3. API permissions → Add → Microsoft Graph → Application permissions → `Mail.Read`
4. Grant admin consent
5. Certificates & secrets → New client secret → copy the value
6. Copy Tenant ID, Client ID, and Client Secret into your `.env` file
