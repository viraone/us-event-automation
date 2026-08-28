# US Event Automation and Gig Dashboard

This project has two coordinated, read-only-for-gigs workflows:

1. `main.py` signs in to US Event Management, reads the calendar, and updates the monthly spreadsheet exports.
2. `app.py` runs a local dashboard that can start that existing updater in the background, then rereads every `US_Event_*.xlsx` workbook.

The dashboard never duplicates the scraper and never claims, accepts, cancels, or modifies gigs. Its only write action is starting the fixed local `main.py` update process. The spreadsheets remain the source of truth.

## Project structure

- `main.py` — existing Playwright scraper and spreadsheet updater
- `app.py` — local Flask dashboard entry point
- `refresh_manager.py` — single-job background runner, rollback protection, and sanitized status
- `dashboard_data.py` — dynamic workbook discovery and normalization
- `templates/` — dashboard page templates
- `static/css/styles.css` — responsive dashboard styling
- `static/js/app.js` — local filtering, sorting, pagination, and recruiter-text generation
- `US_Event_*.xlsx` — monthly source workbooks discovered automatically
- `backups/` — timestamped workbook backups created by the scraper
- `dashboard_state.json` — last successful web refresh metadata (created after the first success)

## Setup

```bash
cd /Users/viradeth/Desktop/US-EVENT-AUTO/us-event-automation
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m playwright install chromium
```

Configure the scraper credentials in `.env`:

```env
US_EVENT_EMAIL=your_email@example.com
US_EVENT_PASSWORD=your_password
```

## Daily workflow

```bash
cd /Users/viradeth/Desktop/US-EVENT-AUTO/us-event-automation
.venv/bin/python app.py
```

Open [http://127.0.0.1:5000](http://127.0.0.1:5000), then click **Refresh Gig Data**. The button starts the existing scraper with the same virtual-environment Python, prevents overlapping runs, and shows the result without requiring a manual page refresh.

To prepare a recruiter message, select any current gigs in the table and click **Generate Recruiter Text**. The dashboard sorts only those selections by date and start time, builds a local message addressed to Adam, and opens it in an editable side panel. **Copy Text** copies the current edited version; the dashboard never sends, claims, or modifies a gig. A successful data refresh clears this temporary selection/message state so it cannot refer to an outdated spreadsheet row.

The scraper still runs with a visible browser, preserves `auth_state.json`, and updates all configured month workbooks. If it fails, the dashboard restores the previous spreadsheet/CSV files and keeps showing the last known-good data. Diagnostic output shown in the browser is limited and sanitized.

Running `.venv/bin/python main.py` directly remains available for command-line debugging, but it is no longer part of the normal daily workflow.

Future files such as `US_Event_October_2026_Gigs.xlsx` are discovered automatically without code changes.

## Run checks

```bash
cd /Users/viradeth/Desktop/US-EVENT-AUTO/us-event-automation
.venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m py_compile main.py app.py dashboard_data.py refresh_manager.py
```
