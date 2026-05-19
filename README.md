**Smart-Job-Scout-AI-Agent**

(**Project Overview**)
- **Purpose**: A scheduled job reporter that searches jobs via the JSearch (RapidAPI) service, writes results to an Excel spreadsheet, and optionally emails a formatted report.

(**Quick Start**)
- **Setup**: create a virtual environment and install dependencies: `pip install -r requirements-jobs.txt`.
- **Run**: execute the batch wrapper: [run_job_reporter.bat](run_job_reporter.bat) which activates the virtualenv Python and runs the reporter.

(**Files**)
- **README**: This file ([README.md](README.md)).
- **Batch runner**: [run_job_reporter.bat](run_job_reporter.bat) — creates a `logs/` folder, logs start/finish, and invokes the script inside `.venv`.
- **Dependencies**: [requirements-jobs.txt](requirements-jobs.txt) — lists Python packages (`httpx`, `openpyxl`, `python-dotenv`).
- **Main script**: [jsearch_job_reporter.py](jsearch_job_reporter.py) — contains `JSearchJobReporter` to call the JSearch API, parse results, export an Excel file, and send email reports.
- **Config loader**: [config_loader.py](config_loader.py) — parses `application.properties` and applies environment variable overrides.
- **Properties**: [application.properties](application.properties) — default configuration values (RapidAPI key, SMTP credentials, email defaults, job query and search settings).

(**Configuration & Secrets**)
- **Primary config**: `application.properties` is read by `config_loader.py`. Environment variables from `ENV_OVERRIDES` (e.g. `RAPIDAPI_KEY`, `SMTP_USER`, `SMTP_PASSWORD`) take precedence for safer secret injection.
- **Important**: `application.properties` in this repo contains real-looking credentials; rotate and move secrets to environment variables or a secure vault before sharing.

(**Behavior Summary**)
- **Search**: `JSearchJobReporter.search_jobs()` queries the RapidAPI JSearch `search-v2` endpoint using `rapidapi.key` and host settings.
- **Parse**: `parse_jobs()` flattens API responses into rows matching the `JOB_FIELDS` used for Excel columns.
- **Export**: `export_to_excel()` writes results to an `.xlsx` file using `openpyxl` and stores it in `output.dir` (default `.`).
- **Email**: `send_email()` composes a plain-text and HTML multipart message, attaches the spreadsheet, and sends through configured SMTP host/port with TLS.

(**How to run manually**)
1. Create and activate a virtual environment inside the repo (optional but recommended).
2. Install dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-jobs.txt
```

3. Provide secrets as env vars or edit `application.properties` (env vars preferred):
- `RAPIDAPI_KEY`, `SMTP_USER`, `SMTP_PASSWORD`, `EMAIL_TO`.

4. Run the reporter directly (from repo root):

```powershell
python jsearch_job_reporter.py
```

or via the batch wrapper on Windows: double-click or run `run_job_reporter.bat`.

(**Notes & Next Steps**)
- **Logging**: `run_job_reporter.bat` writes logs into `logs/` with a timestamped name.
- **Retries / Errors**: HTTP errors raise exceptions in `search_jobs()`; consider adding retry/backoff logic for robustness.
- **Security**: Remove hard-coded secrets from `application.properties` and use environment variables or a secrets manager.

