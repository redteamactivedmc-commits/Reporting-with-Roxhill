"""
File naming utilities following ADMC conventions.

Snapshot  → ADMC_<ClientName>_<MagazineName>_<DateOfPublishing>.docx
Coverage  → ADMC_<ClientName>_<PressReleaseName>_<Date>.docx
Tracker   → ADMC_clients_<ClientName>tracker.xlsx
"""

import re
from datetime import datetime
from media_monitor.config import (
    SNAPSHOT_PREFIX,
    COVERAGE_PREFIX,
    TRACKER_PREFIX,
    SNAPSHOTS_DIR,
    COVERAGE_DIR,
    TRACKER_DIR,
)
import os


def _sanitize(text: str) -> str:
    """Remove or replace characters unsafe for filenames."""
    text = text.strip()
    text = re.sub(r"[^\w\s\-]", "", text)
    text = re.sub(r"\s+", "_", text)
    return text


def _format_date(date_str: str | None) -> str:
    """
    Accept multiple date formats and return DDMMYYYY.
    Falls back to today if parsing fails.
    """
    if not date_str:
        return datetime.today().strftime("%d%m%Y")

    formats = [
        "%d %B %Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%B %d, %Y",
        "%d %b %Y",
        "%d%m%Y",
    ]

    cleaned = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", date_str).strip()

    for fmt in formats:
        try:
            return datetime.strptime(cleaned, fmt).strftime("%d%m%Y")
        except ValueError:
            continue

    return datetime.today().strftime("%d%m%Y")


def snapshot_filename(client: str, magazine: str, date_str: str) -> str:
    """ADMC_<ClientName>_<MagazineName>_<DateOfPublishing>.docx"""
    parts = [
        SNAPSHOT_PREFIX,
        _sanitize(client),
        _sanitize(magazine),
        _format_date(date_str),
    ]
    filename = "_".join(parts) + ".docx"
    return os.path.join(SNAPSHOTS_DIR, filename)


def coverage_filename(client: str, press_release_name: str, date_str: str) -> str:
    """ADMC_<ClientName>_<PressReleaseName>_<Date>.docx"""
    parts = [
        COVERAGE_PREFIX,
        _sanitize(client),
        _sanitize(press_release_name),
        _format_date(date_str),
    ]
    filename = "_".join(parts) + ".docx"
    return os.path.join(COVERAGE_DIR, filename)


def tracker_filename(client: str) -> str:
    """ADMC_clients_<ClientName>tracker.xlsx"""
    filename = f"{TRACKER_PREFIX}_{_sanitize(client)}tracker.xlsx"
    return os.path.join(TRACKER_DIR, filename)
