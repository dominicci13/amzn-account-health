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

### 3. Configure accounts

```bash
cp config/accounts.json.example config/accounts.json
```

Edit `config/accounts.json` with your Amazon account names and their corresponding Excel cell positions.

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
