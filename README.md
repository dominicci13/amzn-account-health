# amzn-account-health-metrics

Weekday automation that scrapes each Amazon Seller Central storefront's **Account Health** (Late Shipment Rate, Pre-Fulfillment Cancel Rate, Valid Tracking Rate), **Prime Performance**, and **Seller Fulfilled Prime** metrics — including pixel-cropped screenshots of each account's speed-distribution chart — writes the results into `AH-Metrics.xlsm`, and emails the workbook.

The script is **one Python file** orchestrating five external surfaces: SeleniumBase (Amazon login + DOM scraping across Account Health / Prime / SFP pages), Pillow + Windows clipboard (chart screenshot crop + paste-into-Excel), xlwings (`AH-Metrics.xlsm` opened hidden, two sheets), Outlook (weekday email), and APScheduler (`Mon-Fri 11:00`).

## Weekday flow

1. **Open + clean** — open `AH-Metrics.xlsm` hidden, call `modUtilities.deleteCharts` to clear any pre-existing chart shapes on the dashboard sheet so the new run pastes onto a clean canvas.
2. **Per-account Account Health + Prime** — log into each account in `AMAZON_URLS`, scrape the Performance Dashboard (Late Shipment Rate, Pre-Fulfillment Cancel Rate, Valid Tracking Rate) and the Prime eligibility page (eligibility status, OTDR, Pre-Fulfillment Cancellation Rate, Valid Tracking Rate). Write into the account's column on the metrics sheet.
3. **SFP per size tier** — for each size tier (Standard / Oversize), navigate to the SFP Performance page, scrape program status, screenshot the speed-distribution chart (crop with Pillow, paste via Windows clipboard onto the dashboard sheet), then scrape the Fulfillment and Supporting metrics tables.
4. **Normalize + save** — call `modUtilities.resizeCharts` to standardize chart dimensions, stamp today's date into both sheets' header cells, save and close the workbook.
5. **Email** — send the refreshed `.xlsm` as an attachment via Outlook.

## Project layout

```
amzn-account-health-metrics/
├── run_amzn_account_health_metrics.py  # entry point (single script)
├── config/
│   ├── accounts.json.example           # Amazon account names + URLs + per-account column/dashboard layout
│   ├── metrics_layout.json.example     # row numbers in the metrics sheet (account-agnostic)
│   └── paths.json.example              # absolute path to AH-Metrics.xlsm
├── vba/
│   └── modUtilities.bas                # canonical VBA source — deleteCharts + resizeCharts
├── screenshots/                        # crash screenshots (gitignored)
├── logs/                               # rotating run logs (gitignored)
├── downloaded_files/                   # Chrome download landing zone (gitignored)
├── output/                             # reserved for future use (gitignored)
├── .env.example
├── requirements.txt
└── README.md
```

## Setup

### 1. Clone and create the venv

```powershell
git clone https://github.com/dominicci13/amzn-account-health-metrics.git
cd amzn-account-health-metrics
py -3.12 -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\pip install git+https://github.com/dominicci13/shared-python-utils.git
```

### 2. Configure

```powershell
copy .env.example .env
copy config\accounts.json.example config\accounts.json
copy config\metrics_layout.json.example config\metrics_layout.json
copy config\paths.json.example config\paths.json
```

Edit each file with real values. All four are gitignored.

`config/accounts.json` has four top-level keys:

- `amazon_account_names` / `amazon_urls` — consumed by `fc_utils.accounts.iter_amazon_accounts()` at import time.
- `account_health_metrics_columns` — maps each account name to its column letter on the metrics sheet (e.g. `"FocusCam": "D"`).
- `account_health_dashboard` — per-account chart anchor cell + speed-table column/row anchors on the dashboard sheet.

`config/metrics_layout.json` holds the row numbers for each metric block on the metrics sheet (account-agnostic — these only change when the workbook itself is restructured).

### 3. VBA module (one-time per workbook)

`AH-Metrics.xlsm` must contain the canonical `modUtilities` from `vba/modUtilities.bas`. Open the workbook in Excel, press **Alt+F11**, insert a module named `modUtilities`, and paste the contents of `vba/modUtilities.bas`. Save the workbook.

### 4. Run

```powershell
.venv\Scripts\python run_amzn_account_health_metrics.py
```

The script prompts "Run now?" — answer **Y** to execute immediately, or **N** to register the APScheduler job and idle until the next **Mon-Fri 11:00** trigger.

## Environment variables

| Variable | Description |
|---|---|
| `AMZN_email` | Amazon Seller Central login email |
| `AMZN_pass` | Amazon Seller Central password |
| `CHROME_USER_DATA_DIR` | Path to the persistent Chrome profile directory used by the bot |
| `SENDER_EMAIL` | Outlook account used to send the weekday email |
| `TO_EMAIL` | Comma-separated recipients |
| `CC_EMAIL` | Comma-separated CC list (optional) |

## License

[MIT](LICENSE)
