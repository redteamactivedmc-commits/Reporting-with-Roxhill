"""
CLI entry point for ADMC Media Monitor.

Usage:
  python -m media_monitor snapshot   --url URL --client CLIENT --magazine MAG [options]
  python -m media_monitor coverage-24h --client CLIENT --campaign CAMPAIGN [options]
  python -m media_monitor tracker update --client CLIENT [options]
"""

import json
import sys
import click
from datetime import datetime
from rich.console import Console

console = Console()


@click.group()
def cli():
    """
    \b
    ADMC Media Monitor -- AI PR Reporter

    Three commands:
      snapshot      -> Capture article screenshot + metadata
      coverage-24h  -> Build 24-hour press coverage report
      tracker       -> Add entries to Excel coverage tracker
    """
    pass


# ── COMMAND 1 -- snapshot ────────────────────────────────────────────────────
@cli.command("snapshot")
@click.option("--url",         required=True,  help="Full URL of the article")
@click.option("--client",      required=True,  help="Client name  e.g. Phosphorus")
@click.option("--magazine",    required=True,  help="Magazine/publication name")
@click.option("--headline",    default="",     help="Article headline")
@click.option("--date",        default="",     help="Publishing date  e.g. '3rd March 2025'")
@click.option("--publication", default="",     help="Publication name (if different from magazine)")
def cmd_snapshot(url, client, magazine, headline, date, publication):
    """Take a snapshot of a media article."""
    from media_monitor.snapshot import create_snapshot
    create_snapshot(
        url=url,
        client=client,
        magazine=magazine,
        headline=headline,
        date_str=date,
        publication=publication or magazine,
    )


# ── COMMAND 2 -- coverage-24h ────────────────────────────────────────────────
@cli.command("coverage-24h")
@click.option("--client",   required=True, help="Client name")
@click.option("--campaign", required=True, help="Press release / campaign name")
@click.option("--date",     default="",    help="Report date (defaults to today)")
@click.option(
    "--items-file",
    default=None,
    help="Path to a JSON file with coverage items (see docs for schema)",
)
def cmd_coverage_24h(client, campaign, date, items_file):
    """Build a 24-hour press coverage report."""
    from media_monitor.coverage_24h import create_coverage_24h

    if items_file:
        with open(items_file, "r", encoding="utf-8") as fh:
            items = json.load(fh)
    else:
        console.print(
            "[yellow]No --items-file provided. "
            "Enter coverage items interactively (Ctrl-C to finish):[/yellow]"
        )
        items = _collect_items_interactively()

    create_coverage_24h(
        client=client,
        campaign=campaign,
        items=items,
        date_str=date,
    )


def _collect_items_interactively() -> list[dict]:
    items = []
    idx   = 1
    while True:
        try:
            console.print(f"\n[cyan]-- Item {idx} --[/cyan]  (press Ctrl-C to stop)")
            pub      = click.prompt("  Publication")
            headline = click.prompt("  Headline")
            date_    = click.prompt("  Date (DD-MM-YYYY)")
            url      = click.prompt("  URL")
            lang     = click.prompt("  Language", default="English")
            po       = click.prompt("  Print/Online", default="Online")
            items.append({
                "publication":  pub,
                "headline":     headline,
                "date":         date_,
                "url":          url,
                "language":     lang,
                "print_online": po,
            })
            idx += 1
        except (KeyboardInterrupt, click.Abort):
            break
    return items


# ── COMMAND 3 -- tracker ─────────────────────────────────────────────────────
@cli.group("tracker")
def tracker_group():
    """Excel Coverage Tracker commands."""
    pass


@tracker_group.command("update")
@click.option("--client",     required=True, help="Client name")
@click.option(
    "--entries-file",
    default=None,
    help="Path to JSON file with tracker entries (see docs for schema)",
)
def cmd_tracker_update(client, entries_file):
    """Add new rows to the Excel tracker."""
    from media_monitor.tracker import update_tracker

    if entries_file:
        with open(entries_file, "r", encoding="utf-8") as fh:
            entries = json.load(fh)
    else:
        console.print(
            "[yellow]No --entries-file provided. "
            "Enter entries interactively (Ctrl-C to finish):[/yellow]"
        )
        entries = _collect_entries_interactively()

    update_tracker(client=client, entries=entries)


def _collect_entries_interactively() -> list[dict]:
    entries = []
    idx     = 1
    while True:
        try:
            console.print(f"\n[cyan]-- Entry {idx} --[/cyan]  (press Ctrl-C to stop)")
            date    = click.prompt("  Date (DD/MM/YYYY)")
            pub     = click.prompt("  Publication")
            hl      = click.prompt("  Headline")
            mtype   = click.prompt("  Media Type", default="IT/ Tech")
            lang    = click.prompt("  Language",   default="English")
            spk     = click.prompt("  Spokesperson")
            src     = click.prompt("  Source (URL)")
            reach   = click.prompt("  Reach",      default="0")
            tier    = click.prompt("  Tier",        default="Tier 1")
            po      = click.prompt("  Print/Online",default="Online")
            pivot   = click.prompt("  For Pivot",   default="")
            entries.append({
                "Date (DD/MM/YYYY)": date,
                "Publication":       pub,
                "Headline":          hl,
                "Media Type":        mtype,
                "Language":          lang,
                "Spokesperson":      spk,
                "Source":            src,
                "Reach":             int(reach) if reach.isdigit() else reach,
                "Tier":              tier,
                "Print/Online":      po,
                "For Pivot":         pivot,
            })
            idx += 1
        except (KeyboardInterrupt, click.Abort):
            break
    return entries


if __name__ == "__main__":
    cli()
