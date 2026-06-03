"""
File naming utilities following ADMC conventions.

Snapshot  -> ADMC_<ClientName>_<MagazineName>_<DateOfPublishing>.docx
Coverage  -> ADMC_<ClientName>_<PressReleaseName>_<Date>.docx
Tracker   -> ADMC_clients_<ClientName>tracker.xlsx
"""

import re
import os
from datetime import datetime
from media_monitor.config import (
    SNAPSHOT_PREFIX,
    COVERAGE_PREFIX,
    TRACKER_PREFIX,
    SNAPSHOTS_DIR,
    COVERAGE_DIR,
    TRACKER_DIR,
)


def _sanitize(text: str) -> str:
    text = text.strip()
    text = re.sub(r"[^\w\s\-]", "", text)
    text = re.sub(r"\s+", "_", text)
    return text


def _format_date(date_str: str | None) -> str:
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
    parts = [
        SNAPSHOT_PREFIX,
        _sanitize(client),
        _sanitize(magazine),
        _format_date(date_str),
    ]
    filename = "_".join(parts) + ".docx"
    return os.path.join(SNAPSHOTS_DIR, filename)


def coverage_filename(client: str, press_release_name: str, date_str: str) -> str:
    parts = [
        COVERAGE_PREFIX,
        _sanitize(client),
        _sanitize(press_release_name),
        _format_date(date_str),
    ]
    filename = "_".join(parts) + ".docx"
    return os.path.join(COVERAGE_DIR, filename)


def tracker_filename(client: str) -> str:
    filename = f"{TRACKER_PREFIX}_{_sanitize(client)}tracker.xlsx"
    return os.path.join(TRACKER_DIR, filename)
