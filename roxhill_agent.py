"""
roxhill_agent.py
────────────────
Main orchestrator for the Roxhill Coverage Agent.

Connects:
  1. OutlookFetcher  → pulls Roxhill alert emails from Outlook
  2. roxhill_parser  → extracts article metadata from email HTML
  3. media_monitor   → generates coverage reports (24h doc, snapshots, tracker)

Run modes:
  python roxhill_agent.py                        # all clients, last 24h, all reports
  python roxhill_agent.py --client uipath        # UiPath only
  python roxhill_agent.py --hours 48             # 48h lookback
  python roxhill_agent.py --output 24h           # 24h coverage doc only
  python roxhill_agent.py --output tracker       # Excel tracker only
  python roxhill_agent.py --output snapshot      # Snapshots only
  python roxhill_agent.py --tier A               # Tier A articles only
  python roxhill_agent.py --dry-run              # Parse and print, no files created

Environment:
  Copy .env.example to .env and fill in your Microsoft credentials.
  Never commit .env to version control.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

# Load .env if python-dotenv is available
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from outlook_fetcher import OutlookFetcher
from roxhill_parser import parse_roxhill_email

# Import the original media monitor commands
# These are the existing modules from the original system
try:
    from media_monitor.coverage_24h import create_coverage_24h
    from media_monitor.snapshot import create_snapshot
    from media_monitor.tracker import update_tracker
    MEDIA_MONITOR_AVAILABLE = True
except ImportError:
    MEDIA_MONITOR_AVAILABLE = False

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Tier filter
# ─────────────────────────────────────────────────────────────────────────────

TIER_ORDER = {"A": 1, "B": 2, "C": 3, "": 9}

def filter_by_tier(articles: list[dict], min_tier: str) -> list[dict]:
    if min_tier == "any":
        return articles
    max_rank = TIER_ORDER.get(min_tier.upper(), 9)
    return [a for a in articles if TIER_ORDER.get(a.get("tier", ""), 9) <= max_rank]


# ─────────────────────────────────────────────────────────────────────────────
# Tracker row builder
# Maps parsed article fields to the ADMC tracker column schema
# ─────────────────────────────────────────────────────────────────────────────

def article_to_tracker_row(article: dict) -> dict:
    tier_label = f"Tier {article['tier']}" if article.get("tier") else ""
    return {
        "Date (DD/MM/YYYY)": article.get("date", ""),
        "Publication":       article.get("publication", ""),
        "Headline":          article.get("headline", ""),
        "Media Type":        "IT / Tech",               # default; override per client if needed
        "Language":          article.get("language", "English"),
        "Spokesperson":      "",                         # not in email; fill manually
        "Source":            article.get("url", ""),
        "Reach":             article.get("reach", ""),
        "Tier":              tier_label,
        "Print/Online":      article.get("media_type", "Online"),
        "For Pivot":         "",
    }


# ─────────────────────────────────────────────────────────────────────────────
# Coverage item builder
# Maps parsed article fields to the 24h coverage report item schema
# ─────────────────────────────────────────────────────────────────────────────

def article_to_coverage_item(article: dict) -> dict:
    return {
        "publication":  article.get("publication", ""),
        "headline":     article.get("headline", ""),
        "date":         article.get("date", ""),
        "url":          article.get("url", ""),
        "language":     article.get("language", "English"),
        "print_online": article.get("media_type", "Online"),
        "screenshot_path": "",  # screenshots require the Selenium backend
    }


# ─────────────────────────────────────────────────────────────────────────────
# Main orchestrator
# ─────────────────────────────────────────────────────────────────────────────

def run_agent(
    lookback_hours: int = 24,
    client_filter: str = "",
    output_type: str = "all",
    min_tier: str = "any",
    dry_run: bool = False,
    export_json: bool = True,
) -> dict:
    """
    Full pipeline:
      Outlook → parse → filter → generate reports

    Returns a summary dict with counts and output file paths.
    """
    logger.info("=" * 60)
    logger.info("Roxhill Coverage Agent starting")
    logger.info(f"  Lookback : {lookback_hours}h")
    logger.info(f"  Client   : {client_filter or 'all'}")
    logger.info(f"  Output   : {output_type}")
    logger.info(f"  Min tier : {min_tier}")
    logger.info(f"  Dry run  : {dry_run}")
    logger.info("=" * 60)

    # ── Step 1: Fetch emails ──────────────────────────────────────────────────
    logger.info("Connecting to Outlook...")
    fetcher = OutlookFetcher()

    client_alias = None
    if client_filter:
        # Build the To: alias from the client name
        # e.g. "uipath" → looks for "uipath@activedmc.com" in To: field
        client_alias = f"{client_filter.lower()}@activedmc.com"

    emails = fetcher.fetch_all_pages(lookback_hours=lookback_hours)
    logger.info(f"Fetched {len(emails)} Roxhill email(s)")

    # ── Step 2: Parse each email ──────────────────────────────────────────────
    all_articles: list[dict] = []

    for email in emails:
        articles = parse_roxhill_email(
            html_body   = email["body_html"],
            to_address  = email["to_address"],
            subject     = email["subject"],
            received_dt = email["received_dt"],
        )

        # Client filter
        if client_filter:
            articles = [
                a for a in articles
                if client_filter.lower() in a["client"].lower()
                or (client_alias and client_alias.lower() in email["to_address"].lower())
            ]

        logger.info(f"  '{email['subject']}' → {len(articles)} article(s)")
        all_articles.extend(articles)

    # ── Step 3: Tier filter ───────────────────────────────────────────────────
    filtered = filter_by_tier(all_articles, min_tier)
    logger.info(f"After tier filter: {len(filtered)}/{len(all_articles)} articles")

    if not filtered:
        logger.warning("No articles passed filters. Exiting.")
        return {"total": 0, "files": []}

    # ── Step 4: Group by client ───────────────────────────────────────────────
    by_client: dict[str, list[dict]] = {}
    for article in filtered:
        client = article["client"]
        by_client.setdefault(client, []).append(article)

    logger.info(f"Clients with coverage: {list(by_client.keys())}")

    # ── Step 5: Export JSON (always, for transparency) ────────────────────────
    output_files: list[str] = []

    if export_json:
        today = datetime.today().strftime("%d%m%Y")
        json_path = f"output/roxhill_parsed_{today}.json"
        os.makedirs("output", exist_ok=True)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(
                {"generated": datetime.now().isoformat(), "articles": filtered},
                f, ensure_ascii=False, indent=2
            )
        logger.info(f"JSON export: {json_path}")
        output_files.append(json_path)

    # ── Step 6: Generate reports ──────────────────────────────────────────────
    if dry_run:
        logger.info("DRY RUN — printing results, no files created")
        for article in filtered:
            print(f"\n  [{article['client']}] [{article['tier']}] [{article['media_type']}]")
            print(f"  {article['headline']}")
            print(f"  {article['publication']} | {article['date']}")
            print(f"  {article['url']}")
        return {"total": len(filtered), "files": output_files}

    if not MEDIA_MONITOR_AVAILABLE:
        logger.warning(
            "media_monitor package not found. "
            "JSON exported. To generate .docx and .xlsx reports, "
            "run this from inside your media_monitor project directory."
        )
        return {"total": len(filtered), "files": output_files}

    today_str = datetime.today().strftime("%-d %B %Y")

    for client, articles in by_client.items():
        logger.info(f"\nGenerating reports for: {client} ({len(articles)} articles)")

        # ── 24h coverage report ───────────────────────────────────────────
        if output_type in ("24h", "all"):
            items = [article_to_coverage_item(a) for a in articles]
            campaign = f"{client} Roxhill Coverage {today_str}"
            path = create_coverage_24h(
                client=client,
                campaign=campaign,
                items=items,
                date_str=today_str,
            )
            output_files.append(path)
            logger.info(f"  24h report: {path}")

        # ── Snapshots (one per article) ───────────────────────────────────
        if output_type in ("snapshot", "all"):
            for article in articles:
                if not article.get("url"):
                    continue
                path = create_snapshot(
                    url         = article["url"],
                    client      = client,
                    magazine    = article.get("publication", "Unknown"),
                    headline    = article.get("headline", ""),
                    date_str    = article.get("date", ""),
                    publication = article.get("publication", ""),
                )
                output_files.append(path)
            logger.info(f"  Snapshots created: {len(articles)}")

        # ── Excel tracker update ──────────────────────────────────────────
        if output_type in ("tracker", "all"):
            entries = [article_to_tracker_row(a) for a in articles]
            path = update_tracker(client=client, entries=entries)
            output_files.append(path)
            logger.info(f"  Tracker updated: {path}")

    logger.info("\n" + "=" * 60)
    logger.info(f"Done. {len(filtered)} articles processed. {len(output_files)} file(s) created.")
    logger.info("=" * 60)

    return {"total": len(filtered), "files": output_files}


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Roxhill Coverage Agent — fetch Roxhill alerts and generate ADMC reports",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python roxhill_agent.py
  python roxhill_agent.py --client uipath --hours 48
  python roxhill_agent.py --output tracker --tier A
  python roxhill_agent.py --dry-run
        """
    )

    parser.add_argument(
        "--client", default="",
        help="Filter by client name or alias key (e.g. uipath, scc, illumio). Default: all clients."
    )
    parser.add_argument(
        "--hours", type=int, default=24,
        help="Lookback window in hours. Default: 24."
    )
    parser.add_argument(
        "--output", default="all",
        choices=["24h", "snapshot", "tracker", "all"],
        help="Which report(s) to generate. Default: all."
    )
    parser.add_argument(
        "--tier", default="any",
        choices=["any", "A", "AB"],
        help="Minimum tier to include. 'A' = Tier A only, 'AB' = A and B. Default: any."
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Parse and print results without creating any files."
    )
    parser.add_argument(
        "--no-json", action="store_true",
        help="Skip JSON export."
    )

    args = parser.parse_args()

    result = run_agent(
        lookback_hours = args.hours,
        client_filter  = args.client,
        output_type    = args.output,
        min_tier       = args.tier,
        dry_run        = args.dry_run,
        export_json    = not args.no_json,
    )

    print(f"\nSummary: {result['total']} articles processed")
    for f in result["files"]:
        print(f"  → {f}")


if __name__ == "__main__":
    main()
