#!/usr/bin/env python3
"""
ADMC Roxhill Coverage Agent - Main Orchestrator

Reads Roxhill alert emails from Outlook (or accepts direct URLs),
parses article metadata, and feeds everything into the ADMC
media_monitor report system.

Usage:
  python roxhill_agent.py                              # All clients, all outputs
  python roxhill_agent.py --client knowbe4             # Single client
  python roxhill_agent.py --hours 48 --output 24h      # 48h lookback, 24h report only
  python roxhill_agent.py --dry-run                    # Parse and print, no files
  python roxhill_agent.py --urls-file links.txt        # Direct URLs instead of email

  python roxhill_agent.py --urls \
    "https://www.securitymea.com/..." \
    "https://www.techpulsemea.com/..."                 # Inline URLs
"""

import os
import sys
import json
import argparse
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from roxhill_agent.roxhill_parser import (
    parse_roxhill_email_html,
    parse_direct_urls,
    resolve_client_name,
    CLIENT_ALIAS_MAP,
)
from roxhill_agent.outlook_fetcher import fetch_roxhill_emails

console = Console()


def _filter_by_client(articles: list[dict], client_name: str) -> list[dict]:
    client_lower = client_name.lower()
    filtered = []
    for art in articles:
        headline_lower = art.get("headline", "").lower()
        pub_lower      = art.get("publication", "").lower()
        url_lower      = art.get("url", "").lower()

        if (client_lower in headline_lower or
            client_lower in pub_lower or
            client_lower in url_lower):
            filtered.append(art)

    return filtered


def _filter_by_tier(articles: list[dict], min_tier: str) -> list[dict]:
    if min_tier.upper() == "ANY":
        return articles
    if min_tier.upper() == "A":
        return [a for a in articles if a.get("tier") == "Tier A"]
    if min_tier.upper() == "AB":
        return [a for a in articles if a.get("tier") in ("Tier A", "Tier B")]
    return articles


def _print_articles_table(articles: list[dict], title: str = "Parsed Articles") -> None:
    table = Table(title=title, show_lines=True)
    table.add_column("#",           style="cyan",   justify="center", width=4)
    table.add_column("Publication", style="yellow",  width=25)
    table.add_column("Headline",    style="white",   width=50)
    table.add_column("Date",        style="white",   width=12)
    table.add_column("Lang",        style="green",   justify="center", width=8)
    table.add_column("Tier",        style="magenta", justify="center", width=8)
    table.add_column("URL",         style="blue",    width=40)

    for i, art in enumerate(articles, 1):
        hl = art.get("headline", "")
        table.add_row(
            str(i),
            art.get("publication", ""),
            hl[:50] + ("..." if len(hl) > 50 else ""),
            art.get("date", ""),
            art.get("language", ""),
            art.get("tier", ""),
            art.get("url", "")[:40] + "...",
        )

    console.print(table)


def _export_json(articles: list[dict], client: str) -> str:
    today = datetime.today().strftime("%Y%m%d")
    filename = f"roxhill_parsed_{client}_{today}.json"
    out_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "output", filename
    )
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(articles, f, indent=2, ensure_ascii=False)
    console.print(f"[green]JSON exported: {out_path}[/green]")
    return out_path


def run_coverage_24h(articles: list[dict], client: str, campaign: str) -> str:
    from media_monitor.coverage_24h import create_coverage_24h
    items = []
    for art in articles:
        items.append({
            "publication":    art["publication"],
            "headline":       art["headline"],
            "date":           art["date"],
            "url":            art["url"],
            "language":       art.get("language", "English"),
            "print_online":   art.get("print_online", "Online"),
            "screenshot_path": "",
        })
    return create_coverage_24h(
        client=client,
        campaign=campaign,
        items=items,
    )


def run_snapshots(articles: list[dict], client: str) -> list[str]:
    from media_monitor.snapshot import create_snapshot
    paths = []
    for art in articles:
        try:
            path = create_snapshot(
                url=art["url"],
                client=client,
                magazine=art["publication"],
                headline=art["headline"],
                date_str=art["date"],
                publication=art["publication"],
            )
            paths.append(path)
        except Exception as e:
            console.print(f"[red]Snapshot failed for {art['url']}: {e}[/red]")
    return paths


def run_tracker_update(articles: list[dict], client: str) -> str:
    from media_monitor.tracker import update_tracker
    entries = []
    for art in articles:
        entries.append({
            "Date (DD/MM/YYYY)": art["date"],
            "Publication":       art["publication"],
            "Headline":          art["headline"],
            "Media Type":        "IT/ Tech",
            "Language":          art.get("language", "English"),
            "Spokesperson":      "",
            "Source":            art["url"],
            "Reach":             "",
            "Tier":              art.get("tier", ""),
            "Print/Online":      art.get("print_online", "Online"),
            "For Pivot":         "",
        })
    return update_tracker(client=client, entries=entries)


def main():
    parser = argparse.ArgumentParser(
        description="ADMC Roxhill Coverage Agent",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python roxhill_agent.py --client knowbe4
  python roxhill_agent.py --hours 48 --output 24h
  python roxhill_agent.py --dry-run
  python roxhill_agent.py --urls "https://example.com/article1" "https://example.com/article2"
  python roxhill_agent.py --urls-file links.txt --client knowbe4
        """,
    )
    parser.add_argument("--client",    default=None,  help="Filter by client alias key")
    parser.add_argument("--hours",     type=int, default=24, help="Lookback window in hours (default: 24)")
    parser.add_argument("--output",    default="all", choices=["24h", "snapshot", "tracker", "all"],
                        help="Which reports to generate")
    parser.add_argument("--tier",      default="any", help="Minimum tier: any / A / AB")
    parser.add_argument("--campaign",  default=None,  help="Campaign/press release name for 24h report")
    parser.add_argument("--dry-run",   action="store_true", help="Parse and print, no file generation")
    parser.add_argument("--no-json",   action="store_true", help="Skip JSON export")
    parser.add_argument("--urls",      nargs="+", default=None, help="Direct article URLs")
    parser.add_argument("--urls-file", default=None, help="File with one URL per line")

    args = parser.parse_args()

    console.print(
        Panel(
            "[bold cyan]ADMC Roxhill Coverage Agent[/bold cyan]\n"
            f"Mode     : {'Direct URLs' if (args.urls or args.urls_file) else 'Outlook Email'}\n"
            f"Client   : [yellow]{args.client or 'All'}[/yellow]\n"
            f"Output   : [yellow]{args.output}[/yellow]\n"
            f"Tier     : [yellow]{args.tier}[/yellow]\n"
            f"Dry-run  : [yellow]{args.dry_run}[/yellow]",
            title="Roxhill Agent",
            border_style="blue",
        )
    )

    # -- Collect articles
    all_articles = []

    if args.urls:
        all_articles = parse_direct_urls(args.urls, client=args.client or "")
    elif args.urls_file:
        with open(args.urls_file, "r") as f:
            urls = [line.strip() for line in f if line.strip() and line.strip().startswith("http")]
        all_articles = parse_direct_urls(urls, client=args.client or "")
    else:
        emails = fetch_roxhill_emails(hours=args.hours)
        for email in emails:
            parsed = parse_roxhill_email_html(email["body_html"])
            all_articles.extend(parsed)

    if not all_articles:
        console.print("[yellow]No articles found.[/yellow]")
        return

    # -- Filter by client
    if args.client:
        client_name = resolve_client_name(args.client)
        all_articles = _filter_by_client(all_articles, client_name)
        if not all_articles:
            console.print(f"[yellow]No articles found for client: {client_name}[/yellow]")
            return
    else:
        client_name = "AllClients"

    # -- Filter by tier
    all_articles = _filter_by_tier(all_articles, args.tier)

    console.print(f"\n[bold]Found {len(all_articles)} article(s)[/bold]")
    _print_articles_table(all_articles)

    # -- Export JSON
    if not args.no_json and not args.dry_run:
        _export_json(all_articles, client_name)

    if args.dry_run:
        console.print("[yellow]Dry-run mode - no files generated.[/yellow]")
        return

    # -- Determine campaign name
    campaign = args.campaign
    if not campaign and all_articles:
        first_headline = all_articles[0].get("headline", "Coverage Report")
        if len(first_headline) > 60:
            first_headline = first_headline[:60]
        campaign = first_headline

    # -- Generate reports
    if args.output in ("24h", "all"):
        console.print("\n[bold cyan]--- 24h Coverage Report ---[/bold cyan]")
        run_coverage_24h(all_articles, client_name, campaign)

    if args.output in ("snapshot", "all"):
        console.print("\n[bold cyan]--- Snapshots ---[/bold cyan]")
        run_snapshots(all_articles, client_name)

    if args.output in ("tracker", "all"):
        console.print("\n[bold cyan]--- Tracker Update ---[/bold cyan]")
        run_tracker_update(all_articles, client_name)

    console.print(
        Panel(
            f"[bold green]All done![/bold green]\n"
            f"Articles processed: {len(all_articles)}\n"
            f"Client: {client_name}\n"
            f"Output: {args.output}",
            title="Complete",
            border_style="green",
        )
    )


if __name__ == "__main__":
    main()
