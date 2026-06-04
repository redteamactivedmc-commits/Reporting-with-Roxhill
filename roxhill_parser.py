"""
roxhill_parser.py
─────────────────
Parses Roxhill Media alert emails (HTML) and returns structured article data.

Roxhill email structure (confirmed from live email):
  - Sender:  monitoring-emails@roxhillmedia.com
  - To:      <clientalias>@activedmc.com  (used to identify client)
  - Subject: "<Client Name> Coverage Update"
  - Body:    HTML with one card per article containing:
               • <h3> headline
               • Time + date tags (inline table cells)
               • Online / Print badge
               • Tier badge: A, B, or C
               • Similarweb reach figure
               • "View article" link wrapped in Roxhill tracking URL
               • Publication name in a small font-size:12px cell
               • Optional "Also appeared in..." syndication note

Roxhill tracking URL format:
  https://email-tracking.pimlico.roxhillmedia.com/CL0/<URL-encoded-real-URL>/1/<hash>/<sig>=<n>

Usage:
  from roxhill_parser import parse_roxhill_email, decode_roxhill_url

  articles = parse_roxhill_email(
      html_body   = email_html_string,
      to_address  = "uipath@activedmc.com",
      subject     = "UiPath Coverage Update",
      received_dt = "2026-06-03T12:00:26.000Z",
  )
"""

from __future__ import annotations

import re
from datetime import datetime, timezone
from typing import Optional
from urllib.parse import unquote
from bs4 import BeautifulSoup


# ─────────────────────────────────────────────────────────────────────────────
# Client alias map
# Add new clients here as you onboard them.
# Key: substring to match in the local part of the To: address (lowercase)
# Value: canonical client name
# ─────────────────────────────────────────────────────────────────────────────
CLIENT_ALIAS_MAP: dict[str, str] = {
    "uipath":        "UiPath",
    "illumio":       "Illumio",
    "denodo":        "Denodo",
    "phosphorus":    "Phosphorus Cybersecurity",
    "knowbe4":       "KnowBe4",
    "netscout":      "NETSCOUT",
    "dynatrace":     "Dynatrace",
    "qlik":          "Qlik",
    "heidrick":      "Heidrick & Struggles",
    "levelinfinite": "Level Infinite",
    "level":         "Level Infinite",
}

ROXHILL_SENDER = "monitoring-emails@roxhillmedia.com"


# ─────────────────────────────────────────────────────────────────────────────
# URL decoder
# ─────────────────────────────────────────────────────────────────────────────

def decode_roxhill_url(tracking_url: str) -> str:
    """
    Extract the real article URL from a Roxhill tracking URL.

    Roxhill wraps every article link in:
      https://email-tracking.pimlico.roxhillmedia.com/CL0/<encoded-url>/1/<hash>=<n>

    The encoded portion uses standard URL encoding (colons → %3A, slashes → %2F).
    """
    if not tracking_url or "email-tracking.pimlico.roxhillmedia.com" not in tracking_url:
        return tracking_url

    # Pattern: /CL0/<encoded-url>/<rest>
    match = re.search(r"/CL0/(https?:[^/]+)/", tracking_url)
    if match:
        return unquote(match.group(1))

    return tracking_url


# ─────────────────────────────────────────────────────────────────────────────
# Client resolver
# ─────────────────────────────────────────────────────────────────────────────

def resolve_client(to_address: str, subject: str) -> str:
    """
    Identify the client from the To: address or email subject.

    Priority:
      1. Match local part of To: address against CLIENT_ALIAS_MAP
      2. Match subject line against CLIENT_ALIAS_MAP keys
      3. Strip "Coverage Update" from subject and use remainder as fallback
    """
    if to_address:
        local = to_address.split("@")[0].lower()
        for key, name in CLIENT_ALIAS_MAP.items():
            if key in local:
                return name

    if subject:
        sub_lower = subject.lower()
        for key, name in CLIENT_ALIAS_MAP.items():
            if key in sub_lower:
                return name
        # Fallback: strip boilerplate from subject
        cleaned = re.sub(r"coverage update", "", subject, flags=re.IGNORECASE).strip(" –-|")
        if cleaned:
            return cleaned

    return "Unknown Client"


# ─────────────────────────────────────────────────────────────────────────────
# Date formatter
# ─────────────────────────────────────────────────────────────────────────────

def _format_date(iso_string: Optional[str]) -> str:
    """Return DD/MM/YYYY from an ISO datetime string. Falls back to today."""
    if not iso_string:
        return datetime.now().strftime("%d/%m/%Y")
    try:
        dt = datetime.fromisoformat(iso_string.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return datetime.now().strftime("%d/%m/%Y")


# ─────────────────────────────────────────────────────────────────────────────
# Article block detector
# ─────────────────────────────────────────────────────────────────────────────

def _extract_publication(block_html: str) -> str:
    """
    Find the publication name within an article block.
    Roxhill puts publication names in small 12px font-size cells
    that appear after the article image column.
    """
    soup = BeautifulSoup(block_html, "lxml")
    candidates = []

    for td in soup.find_all("td"):
        style = td.get("style", "")
        if "font-size:12px" in style or "font-size: 12px" in style:
            text = td.get_text(strip=True)
            # Exclude timestamps, reach figures, and known non-pub strings
            if (
                text
                and len(text) > 2
                and not re.match(r"^\d{1,2}:\d{2}", text)          # timestamps
                and not re.match(r"^Similarweb", text, re.I)        # reach
                and not re.match(r"^Also appeared", text, re.I)     # syndication
                and not re.match(r"^\d{2}\s+\w+\s+\d{4}$", text)   # dates
                and "roxhill" not in text.lower()
            ):
                candidates.append(text)

    # The first qualifying candidate is almost always the publication name
    return candidates[0] if candidates else ""


def _extract_tier(block_html: str) -> str:
    """
    Roxhill renders tier as a pill badge containing a single letter: A, B, or C.
    It appears in a 10px font-size cell with border-radius styling.
    """
    match = re.search(
        r'font-size:10px[^>]*>\s*([A-C])\s*<',
        block_html
    )
    if match:
        return match.group(1)

    # Fallback: look for tier pill pattern without font-size attribute
    match2 = re.search(r'padding:2px 4px[^>]*>\s*([A-C])\s*<', block_html)
    if match2:
        return match2.group(1)

    return ""


def _extract_reach(block_html: str) -> str:
    """Extract Similarweb reach figure, e.g. '68k', '2.4k'."""
    match = re.search(r"Similarweb:\s*([\d.]+k?)\s*\(", block_html, re.I)
    return match.group(1) if match else ""


def _extract_media_type(block_html: str) -> str:
    """Return 'Print' or 'Online' from the badge in the article block."""
    soup = BeautifulSoup(block_html, "lxml")
    for td in soup.find_all("td"):
        text = td.get_text(strip=True)
        if text == "Print":
            return "Print"
        if text == "Online":
            return "Online"
    return "Online"  # default


def _extract_time_and_date(block_html: str) -> tuple[str, str]:
    """
    Roxhill shows: '11:43 UK Time' and 'Wed 03 Jun 2026' as separate inline cells.
    Returns (time_str, date_str) where date_str is DD/MM/YYYY.
    """
    time_match = re.search(r"(\d{1,2}:\d{2})\s+UK Time", block_html)
    date_match = re.search(r"\w{3}\s+(\d{2})\s+(\w{3})\s+(\d{4})", block_html)

    time_str = time_match.group(1) if time_match else ""
    date_str = ""
    if date_match:
        try:
            date_str = datetime.strptime(
                f"{date_match.group(1)} {date_match.group(2)} {date_match.group(3)}",
                "%d %b %Y"
            ).strftime("%d/%m/%Y")
        except ValueError:
            date_str = ""

    return time_str, date_str


def _extract_view_link(block_html: str) -> str:
    """
    Find the 'View article' anchor in the block and decode it.
    All Roxhill article links pass through their tracker.
    """
    soup = BeautifulSoup(block_html, "lxml")
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True).lower()
        href = a["href"]
        if "view article" in text or "email-tracking.pimlico.roxhillmedia.com" in href:
            return decode_roxhill_url(href)
    return ""


def _extract_syndication_note(block_html: str) -> list[str]:
    """
    Extract the 'Also appeared in...' publications list if present.
    These are additional placements with no individual links.
    """
    match = re.search(
        r"Also appeared in\s*([^\n<]+)",
        block_html,
        re.I
    )
    if not match:
        return []
    raw = match.group(1).strip()
    # Split on commas and clean up
    return [p.strip() for p in raw.split(",") if p.strip()]


# ─────────────────────────────────────────────────────────────────────────────
# Main parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_roxhill_email(
    html_body: str,
    to_address: str = "",
    subject: str = "",
    received_dt: str = "",
) -> list[dict]:
    """
    Parse a single Roxhill alert email body and return a list of article dicts.

    Each article dict has:
        client          str   Client name (resolved from To: address or subject)
        headline        str   Article headline
        publication     str   Publication / outlet name
        url             str   Decoded article URL
        date            str   DD/MM/YYYY
        time_uk         str   HH:MM (UK time as shown in the email)
        tier            str   A, B, C, or ""
        media_type      str   "Online" or "Print"
        language        str   "English" (Arabic detection can be added)
        reach           str   Similarweb figure e.g. "68k"
        also_in         list  Syndication outlets (no links)
        email_subject   str   Original email subject

    Returns empty list if html_body is empty or no articles are found.
    """
    if not html_body:
        return []

    client = resolve_client(to_address, subject)
    fallback_date = _format_date(received_dt)
    soup = BeautifulSoup(html_body, "lxml")
    articles = []
    seen_headlines: set[str] = set()

    # Each article is anchored by its <h3> headline element
    for h3 in soup.find_all("h3"):
        headline = h3.get_text(strip=True)
        if not headline or headline in seen_headlines:
            continue
        seen_headlines.add(headline)

        # Walk up to find the containing card block
        # Roxhill wraps each article in nested <div> or <table> blocks
        block_el = h3
        for _ in range(8):
            parent = block_el.parent
            if parent is None:
                break
            block_el = parent
            tag = getattr(block_el, "name", "")
            if tag in ("table", "div") and len(str(block_el)) > 400:
                break

        block_html = str(block_el)

        # Extract all fields
        url             = _extract_view_link(block_html)
        publication     = _extract_publication(block_html)
        tier            = _extract_tier(block_html)
        reach           = _extract_reach(block_html)
        media_type      = _extract_media_type(block_html)
        time_uk, date   = _extract_time_and_date(block_html)
        also_in         = _extract_syndication_note(block_html)

        # Detect Arabic (basic heuristic — Roxhill rarely sends Arabic URLs but it happens)
        language = "Arabic" if url and any(
            dom in url for dom in [
                "alriyadh.com", "alarabiya.net", "asharq.com",
                "zawya.com/ar", "albayan.ae", "alkhaleej.ae",
                "alwatan.com", "okaz.com.sa", "sabq.org"
            ]
        ) else "English"

        articles.append({
            "client":       client,
            "headline":     headline,
            "publication":  publication,
            "url":          url,
            "date":         date or fallback_date,
            "time_uk":      time_uk,
            "tier":         tier,
            "media_type":   media_type,
            "language":     language,
            "reach":        reach,
            "also_in":      also_in,
            "email_subject": subject,
        })

    return articles
