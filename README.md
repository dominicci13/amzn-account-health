# amzn-account-health-metrics

Scrapes Amazon Seller Central account health metrics (policy compliance, voice of the customer, performance ratings) across multiple seller accounts and writes the results into a shared Excel workbook. Runs on a weekday schedule via APScheduler.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
pip install git+https://github.com/dominicci13/shared-python-utils.git
```

### 2. Configure environment

```bash
cp .env.example .env
```

Edit `.env` with your credentials.

### 3. Configure accounts and workbook layout

```bash
cp config/accounts.json.example config/accounts.json
cp config/metrics_layout.json.example config/metrics_layout.json
cp config/paths.json.example config/paths.json
```

- `config/accounts.json` — Amazon account display names, Seller Central URLs, and per-account Excel column/dashboard cell positions. Loaded by `fc_utils.accounts` (shared keys: `amazon_account_names`, `amazon_urls`, `ebay_profiles`) and by this script (script-specific keys: `account_health_metrics_columns`, `account_health_dashboard`).
- `config/metrics_layout.json` — row numbers in the metrics sheet (account-agnostic).
- `config/paths.json` — absolute path to the `AH-Metrics.xlsm` workbook.

## Run

```bash
python run_amzn_account_health_metrics.py
```

The script runs automatically at 11:00 Mon–Fri via APScheduler.

## Environment Variables

| Variable | Description |
|---|---|
| `AMZN_email` | Amazon Seller Central login email |
| `AMZN_pass` | Amazon Seller Central password |
| `CHROME_USER_DATA_DIR` | Path to Chrome automation profile directory |
| `SENDER_EMAIL` | Outlook account used to send the report email |
| `TO_EMAIL` | Comma-separated list of recipient email addresses |
| `CC_EMAIL` | Comma-separated list of CC email addresses |
