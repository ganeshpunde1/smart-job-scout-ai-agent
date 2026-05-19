"""
Fetch job listings from JSearch (RapidAPI), export to Excel, and email the report.
Configuration: application.properties (env vars override file values).
"""

from __future__ import annotations

import html as html_lib
import smtplib
from datetime import datetime, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

import certifi
import httpx
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from config_loader import get_bool, get_int, load_application_config

# Columns exported to Excel.
JOB_FIELDS = (
    ("job_id", "Job ID"),
    ("job_title", "Job Title"),
    ("employer_name", "Employer"),
    ("job_city", "City"),
    ("job_state", "State"),
    ("job_country", "Country"),
    ("job_employment_type", "Employment Type"),
    ("job_description", "Description"),
    ("job_posted_at_datetime_utc", "Posted (UTC)"),
    ("job_min_salary", "Min Salary"),
    ("job_max_salary", "Max Salary"),
    ("job_salary_currency", "Salary Currency"),
    ("job_salary_period", "Salary Period"),
    ("job_publisher", "Publisher"),
    ("job_apply_link", "Apply Link"),
)

JOB_NUMBER_EMOJIS = (
    "1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"
)

DEFAULT_EMAIL_GREETING = "Hi Job Seeker,"
DEFAULT_EMAIL_OPENING = "I hope you are doing well."
DEFAULT_EMAIL_INTRO = (
    "Please find below the latest AI/GenAI job search report curated for Lead, "
    "Senior, and Principal-level opportunities across AI/ML, LLMs, RAG, AWS "
    "Bedrock, Python, and Full Stack Engineering roles."
)
DEFAULT_EMAIL_SIGNER = "Ganesh Punde"

class JSearchJobReporter:
    """Search jobs via JSearch, write Excel, and email results."""

    def __init__(
        self,
        config: dict[str, str] | None = None,
        *,
        rapidapi_key: str | None = None,
        smtp_host: str | None = None,
        smtp_port: int | None = None,
        smtp_user: str | None = None,
        smtp_password: str | None = None,
        email_from: str | None = None,
        email_to: str | list[str] | None = None,
        output_dir: str | Path | None = None,
        timeout: float | None = None,
    ) -> None:
        cfg = config or load_application_config()

        self.rapidapi_key = rapidapi_key or cfg.get("rapidapi.key", "")
        self.rapidapi_host = cfg.get("rapidapi.host", "jsearch.p.rapidapi.com")
        self.base_url = cfg.get(
            "rapidapi.base.url",
            "https://jsearch.p.rapidapi.com/search-v2",
        )
        self.smtp_host = smtp_host or cfg.get("smtp.host", "smtp.gmail.com")
        self.smtp_port = smtp_port or get_int(cfg, "smtp.port", 587)
        self.smtp_user = smtp_user or cfg.get("smtp.user", "")
        self.smtp_password = smtp_password or cfg.get("smtp.password", "")
        self.email_from = email_from or cfg.get("email.from", self.smtp_user)
        self.email_to = self._normalize_recipients(
            email_to or cfg.get("email.to", "")
        )
        self.output_dir = Path(output_dir or cfg.get("output.dir", "."))
        self.timeout = timeout if timeout is not None else float(
            cfg.get("http.timeout", "30")
        )
        self.ssl_verify = get_bool(cfg, "ssl.verify", default=False)
        self.email_greeting = cfg.get("email.greeting", DEFAULT_EMAIL_GREETING)
        self.email_opening = cfg.get("email.opening", DEFAULT_EMAIL_OPENING)
        self.email_intro = cfg.get("email.intro", DEFAULT_EMAIL_INTRO)
        self.email_signer = cfg.get("email.signer", DEFAULT_EMAIL_SIGNER)

        if not self.rapidapi_key:
            raise ValueError(
                "RapidAPI key required. Set rapidapi.key in application.properties."
            )

    @staticmethod
    def _format_posted_date(posted: str | None) -> str | None:
        if not posted:
            return None
        try:
            normalized = posted.replace("Z", "+00:00")
            dt = datetime.fromisoformat(normalized)
            return f"{dt.strftime('%B')} {dt.day}, {dt.year}"
        except ValueError:
            return posted

    @staticmethod
    def _format_salary(job: dict[str, Any]) -> tuple[str | None, str | None, str | None]:
        """Return (salary_range, min_salary, max_salary) display strings."""
        currency = job.get("job_salary_currency") or ""
        period = job.get("job_salary_period") or ""
        suffix = ""
        if currency and period:
            suffix = f" {currency} / {period}"
        elif currency:
            suffix = f" {currency}"
        elif period:
            suffix = f" / {period}"

        min_val = job.get("job_min_salary")
        max_val = job.get("job_max_salary")

        min_str = f"{min_val}{suffix}" if min_val is not None else None
        max_str = f"{max_val}{suffix}" if max_val is not None else None

        if min_val is not None and max_val is not None:
            salary_str = f"{min_val} - {max_val}{suffix}"
        elif min_val is not None:
            salary_str = min_str
        elif max_val is not None:
            salary_str = max_str
        else:
            salary_str = None

        return salary_str, min_str, max_str

    @staticmethod
    def _parse_job_title(title: str) -> tuple[str, str | None]:
        """Split job title and skills noted in parentheses."""
        if "(" in title and ")" in title:
            main_title = title[: title.index("(")].strip()
            skills_raw = title[title.index("(") + 1 : title.rindex(")")].strip()
            skills = " • ".join(
                part.strip()
                for part in skills_raw.replace("/", ",").split(",")
                if part.strip()
            )
            return main_title or title, skills or None
        return title, None

    @staticmethod
    def _is_remote_job(job: dict[str, Any], title: str) -> bool:
        if "remote" in title.lower():
            return True
        employment = str(job.get("job_employment_type") or "").upper()
        return "REMOTE" in employment

    @staticmethod
    def _format_location(job: dict[str, Any]) -> str | None:
        city = job.get("job_city")
        state = job.get("job_state")
        country = job.get("job_country")
        if city and state:
            return f"{city}, {state}"
        if city:
            return str(city)
        if state:
            return str(state)
        if country and country not in ("US", "USA"):
            return str(country)
        return None

    @staticmethod
    def _normalize_recipients(value: str | list[str]) -> list[str]:
        if isinstance(value, list):
            return [e.strip() for e in value if e and e.strip()]
        return [e.strip() for e in str(value).split(",") if e.strip()]

    def _headers(self) -> dict[str, str]:
        return {
            "Content-Type": "application/json",
            "x-rapidapi-host": self.rapidapi_host,
            "x-rapidapi-key": self.rapidapi_key,
        }

    def search_jobs(
        self,
        query: str,
        *,
        num_pages: int = 1,
        country: str = "us",
        date_posted: str = "all",
        employment_types: str | None = None,
        page: int = 1,
        **extra_params: Any,
    ) -> dict[str, Any]:
        """GET /search-v2 and return parsed JSON."""
        params: dict[str, Any] = {
            "query": query,
            "num_pages": num_pages,
            "country": country,
            "date_posted": date_posted,
            "page": page,
            **extra_params,
        }
        if employment_types:
            params["employment_types"] = employment_types
        verify: bool | str = certifi.where() if self.ssl_verify else False
        with httpx.Client(timeout=self.timeout, verify=verify) as client:
            response = client.get(
                self.base_url,
                params=params,
                headers=self._headers(),
            )
            response.raise_for_status()
            return response.json()

    def parse_jobs(self, api_response: dict[str, Any]) -> list[dict[str, Any]]:
        """Extract flat job rows from the API response."""
        data = api_response.get("data")
        if isinstance(data, dict):
            raw_jobs = data.get("jobs") or []
        elif isinstance(data, list):
            raw_jobs = data
        else:
            raw_jobs = []
        rows: list[dict[str, Any]] = []

        for seq, job in enumerate(raw_jobs, start=1):
            if not isinstance(job, dict):
                continue
            row = {key: job.get(key) for key, _ in JOB_FIELDS}
            row["job_id"] = seq
            if not row.get("job_apply_link") and job.get("apply_options"):
                options = job["apply_options"]
                if isinstance(options, list) and options:
                    first = options[0]
                    if isinstance(first, dict):
                        row["job_apply_link"] = first.get("apply_link")
            rows.append(row)

        return rows

    def export_to_excel(
        self,
        jobs: list[dict[str, Any]],
        filename: str | None = None,
    ) -> Path:
        """Write jobs to an .xlsx file; returns the file path."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        if not filename:
            stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
            filename = f"job_search_{stamp}.xlsx"
        path = self.output_dir / filename

        wb = Workbook()
        ws = wb.active
        ws.title = "Jobs"

        headers = [label for _, label in JOB_FIELDS]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)

        for job in jobs:
            ws.append([job.get(key) for key, _ in JOB_FIELDS])

        for col_idx, (_, label) in enumerate(JOB_FIELDS, start=1):
            letter = get_column_letter(col_idx)
            max_len = max(len(label), 12)
            for row in ws.iter_rows(
                min_row=2,
                min_col=col_idx,
                max_col=col_idx,
                max_row=ws.max_row,
            ):
                value = row[0].value
                if value is not None:
                    max_len = max(max_len, min(len(str(value)), 60))
            ws.column_dimensions[letter].width = max_len + 2

        wb.save(path)
        return path

    def _get_job_block(
        self, job: dict[str, Any], index: int
    ) -> tuple[str, str, list[str]]:
        """Return emoji, title, and detail lines for one job listing."""
        emoji = (
            JOB_NUMBER_EMOJIS[index]
            if index < len(JOB_NUMBER_EMOJIS)
            else f"{index + 1}."
        )
        raw_title = job.get("job_title") or "N/A"
        title, skills = self._parse_job_title(raw_title)
        employer = job.get("employer_name") or "N/A"
        location = self._format_location(job)
        posted = self._format_posted_date(job.get("job_posted_at_datetime_utc"))
        salary, _, _ = self._format_salary(job)
        apply_link = job.get("job_apply_link") or "N/A"
        is_remote = self._is_remote_job(job, raw_title)

        details = [f"🏢 {employer}"]
        if skills:
            details.append(f"🛠️ {skills}")
        if is_remote:
            details.append("📍 Remote")
        elif location:
            details.append(f"📍 {location}")
        if posted:
            details.append(f"📅 Posted: {posted}")
        details.append(f"💰 Salary: {salary or 'N/A'}")
        details.append("🔗 Apply:"+apply_link)
        return emoji, title, details

    def _build_email_body(
        self,
        jobs: list[dict[str, Any]],
        query: str,
        excel_path: Path,
    ) -> str:
        report_date = datetime.now().strftime("%B %d, %Y")
        divider = "━" * 22

        lines = [
            self.email_greeting,
            "",
            self.email_opening,
            "",
            self.email_intro,
            "",
            divider,
            "📌 Report Overview",
            divider,
            "",
            f"✨ Total Opportunities Found: {len(jobs)}",
            f"📅 Generated On: {report_date}",
            f"📎 Attachment: {excel_path.name}",
            "",
            divider,
            "🔥 Featured AI/GenAI Opportunities",
            divider,
            "",
        ]

        featured = jobs[:10]
        for index, job in enumerate(featured):
            emoji, title, details = self._get_job_block(job, index)
            lines.append(f"### {emoji} {title}")
            lines.extend(details)
            if index < len(featured) - 1:
                lines.append("---")

        if len(jobs) > 10:
            lines.append(
                f"... and {len(jobs) - 10} more in the attached spreadsheet."
            )

        lines.extend(
            [
                "",
                "Please review the attached Excel report for complete details "
                "and application links.",
                "",
                "Best Regards,",
                self.email_signer,
            ]
        )
        return "\n".join(lines)

    def _build_email_html(
        self,
        jobs: list[dict[str, Any]],
        query: str,
        excel_path: Path,
    ) -> str:
        report_date = datetime.now().strftime("%B %d, %Y")
        divider = "━" * 22
        esc = html_lib.escape

        parts = [
            "<html><body style=\"font-family:Arial,sans-serif;"
            "font-size:14px;line-height:1.45;color:#222;\">",
            f"<p>{esc(self.email_greeting)}</p>",
            f"<p>{esc(self.email_opening)}</p>",
            f"<p>{esc(self.email_intro)}</p>",
            f"<p style=\"white-space:pre-line;\">{esc(divider)}\n"
            f"📌 Report Overview\n{esc(divider)}</p>",
            "<p>"
            f"✨ Total Opportunities Found: {len(jobs)}<br>"
            f"📅 Generated On: {esc(report_date)}<br>"
            f"📎 Attachment: {esc(excel_path.name)}"
            "</p>",
            f"<p style=\"white-space:pre-line;\">{esc(divider)}\n"
            f"🔥 Featured AI/GenAI Opportunities\n{esc(divider)}</p>",
        ]

        featured = jobs[:10]
        for index, job in enumerate(featured):
            emoji, title, details = self._get_job_block(job, index)
            detail_html = "<br>".join(esc(line) for line in details)
            parts.append(
                "<p style=\"margin:0 0 10px 0;\">"
                f"<b>{esc(emoji)} {esc(title)}</b><br>{detail_html}"
                "</p>"
            )
            if index < len(featured) - 1:
                parts.append(
                    "<hr style=\"border:0;border-top:1px solid #ddd;"
                    "margin:8px 0;\">"
                )

        if len(jobs) > 10:
            parts.append(
                f"<p>... and {len(jobs) - 10} more in the attached spreadsheet.</p>"
            )

        parts.extend(
            [
                "<p>Please review the attached Excel report for complete details "
                "and application links.</p>",
                f"<p>Best Regards,<br>{esc(self.email_signer)}</p>",
                "</body></html>",
            ]
        )
        return "".join(parts)

    def send_email(
        self,
        jobs: list[dict[str, Any]],
        excel_path: Path,
        query: str,
        *,
        date_posted: str = "all",
        subject: str | None = None,
        recipients: str | list[str] | None = None,
    ) -> None:
        """Send email with job summary and Excel attachment."""
        to_addrs = self._normalize_recipients(recipients or self.email_to)
        if not to_addrs:
            raise ValueError(
                "Email recipient required. Set email.to in application.properties."
            )
        if not self.smtp_user or not self.smtp_password:
            raise ValueError(
                "SMTP credentials required. Set smtp.user and smtp.password "
                "in application.properties."
            )
        if not self.email_from:
            raise ValueError(
                "Sender required. Set email.from or smtp.user in application.properties."
            )

        subject = subject or f"Job Search Report : {date_posted} ({len(jobs)} jobs)"
        plain_body = self._build_email_body(jobs, query, excel_path)
        html_body = self._build_email_html(jobs, query, excel_path)

        msg = MIMEMultipart()
        msg["From"] = self.email_from
        msg["To"] = ", ".join(to_addrs)
        msg["Subject"] = subject

        alternative = MIMEMultipart("alternative")
        alternative.attach(MIMEText(plain_body, "plain"))
        alternative.attach(MIMEText(html_body, "html"))
        msg.attach(alternative)

        with open(excel_path, "rb") as f:
            part = MIMEBase(
                "application",
                "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{excel_path.name}"',
        )
        msg.attach(part)

        with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
            server.starttls()
            server.login(self.smtp_user, self.smtp_password)
            server.sendmail(self.email_from, to_addrs, msg.as_string())

    def run(
        self,
        query: str,
        *,
        num_pages: int = 1,
        country: str = "us",
        date_posted: str = "all",
        employment_types: str | None = None,
        excel_filename: str | None = None,
        send_mail: bool = True,
        email_subject: str | None = None,
        recipients: str | list[str] | None = None,
    ) -> dict[str, Any]:
        """
        Full pipeline: search -> parse -> Excel -> optional email.

        Returns dict with api_response, jobs, and excel_path.
        """
        api_response = self.search_jobs(
            query,
            num_pages=num_pages,
            country=country,
            date_posted=date_posted,
            employment_types=employment_types,
        )
        jobs = self.parse_jobs(api_response)
        excel_path = self.export_to_excel(jobs, filename=excel_filename)

        result: dict[str, Any] = {
            "status": api_response.get("status"),
            "request_id": api_response.get("request_id"),
            "job_count": len(jobs),
            "jobs": jobs,
            "excel_path": str(excel_path),
        }

        if send_mail:
            if not self.smtp_user or not self.smtp_password:
                result["email_sent"] = False
                result["email_error"] = (
                    "SMTP credentials missing. Set smtp.user and smtp.password "
                    "in application.properties."
                )
            else:
                self.send_email(
                    jobs,
                    excel_path,
                    query,
                    date_posted=date_posted,
                    subject=email_subject,
                    recipients=recipients,
                )
                result["email_sent"] = True
                result["email_to"] = self._normalize_recipients(
                    recipients or self.email_to
                )

        return result


if __name__ == "__main__":
    config = load_application_config()
    reporter = JSearchJobReporter(config)

    query = config.get("job.query", "").strip()
    if not query:
        raise ValueError("job.query is required in application.properties")

    outcome = reporter.run(
        query,
        num_pages=get_int(config, "job.num.pages", 1),
        country=config.get("job.country", "us"),
        date_posted=config.get("job.date.posted", "3days"),
        employment_types=config.get("job.employment.types") or None,
        send_mail=get_bool(config, "job.send.mail", default=True),
    )
    print(f"Found {outcome['job_count']} jobs -> {outcome['excel_path']}")
    if outcome.get("email_sent"):
        print(f"Email sent to: {', '.join(outcome['email_to'])}")
    elif outcome.get("email_error"):
        print(f"Email not sent: {outcome['email_error']}")
