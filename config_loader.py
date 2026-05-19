"""Load settings from application.properties (env vars override file values)."""

from __future__ import annotations

import os
from pathlib import Path

PROPERTIES_FILE = Path(__file__).resolve().parent / "application.properties"

# Environment variable -> property key
ENV_OVERRIDES = {
    "RAPIDAPI_KEY": "rapidapi.key",
    "SMTP_HOST": "smtp.host",
    "SMTP_PORT": "smtp.port",
    "SMTP_USER": "smtp.user",
    "SMTP_PASSWORD": "smtp.password",
    "EMAIL_FROM": "email.from",
    "EMAIL_TO": "email.to",
    "OUTPUT_DIR": "output.dir",
    "SSL_VERIFY": "ssl.verify",
    "HTTP_TIMEOUT": "http.timeout",
    "JOB_QUERY": "job.query",
    "JOB_NUM_PAGES": "job.num.pages",
    "JOB_COUNTRY": "job.country",
    "JOB_DATE_POSTED": "job.date.posted",
    "JOB_EMPLOYMENT_TYPES": "job.employment.types",
    "JOB_SEND_MAIL": "job.send.mail",
}


def load_properties(path: Path | None = None) -> dict[str, str]:
    """Parse a Java-style .properties file."""
    path = path or PROPERTIES_FILE
    props: dict[str, str] = {}

    if not path.is_file():
        return props

    with path.open(encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#") or line.startswith("!"):
                continue
            if "=" not in line:
                continue
            key, _, value = line.partition("=")
            props[key.strip()] = value.strip()

    return props


def load_application_config(path: Path | None = None) -> dict[str, str]:
    """Load application.properties and apply environment overrides."""
    config = load_properties(path)
    for env_name, prop_key in ENV_OVERRIDES.items():
        value = os.environ.get(env_name)
        if value is not None and value != "":
            config[prop_key] = value
    return config


def get_bool(config: dict[str, str], key: str, default: bool = False) -> bool:
    value = config.get(key, str(default)).lower()
    return value in ("1", "true", "yes", "on")


def get_int(config: dict[str, str], key: str, default: int) -> int:
    return int(config.get(key, str(default)))
