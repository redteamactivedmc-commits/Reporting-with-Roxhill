"""
ADMC Media Monitor — Anthropic AI Client
Orchestrates the three reporting commands via the Anthropic API:
  1. snapshot     — capture article screenshot + metadata
  2. coverage-24h — build 24-hour press coverage report
  3. tracker      — add entries to Excel coverage tracker
"""

import os
import json
import sys
import click
from datetime import datetime
from rich.console import Console
from rich.panel import Panel
import anthropic

console = Console()

# Hardcoded KnowBe4 links from the original brief
KNOWBE4_LINKS = [
    "https://www.securitymea.com/2026/03/25/knowbe4-launches-phish-alert-button-for-microsoft-teams/",
    "https://www.techpulsemea.com/knowbe4-brings-one-click-phishing-reporting-to-microsoft-teams/",
    "https://meatechwatch.com/knowbe4-expands-critical-security-defenses-with-phish-alert-button-for-microsoft-teams/",
    "https://techafricanews.com/2026/03/24/knowbe4-launches-phish-alert-button-for-microsoft-teams-to-tackle-rising-chat-based-cyber-threats/",
    "https://gulftech-news.com/knowbe4-expands-critical-security-defenses-with-phish-alert-button-for-microsoft-teams/",
]

SYSTEM_PROMPT = """You are an ADMC Media Monitor AI agent for PR reporting.

You help with three commands:
1. create_snapshot  — Take a snapshot of a media article (screenshot + metadata Word doc)
2. create_coverage  — Build a 24-hour press coverage report (Word doc)
3. update_tracker   — Add entries to the Excel coverage tracker

File naming conventions (strictly enforced):
- Snapshot  : ADMC_<ClientName>_<MagazineName>_<DateOfPublishing>.docx
- Coverage  : ADMC_<ClientName>_<PressReleaseName>_<Date>.docx
- Tracker   : ADMC_clients_<ClientName>tracker.xlsx

Every document must have the client logo (top-left) and Active/ADMC logo (top-right).

When given URLs, use web_search to look up article details (headline, publication name, date).
Then call the appropriate tool with the gathered data.

Always respond with a clear summary of what was created and the output file path."""


def _build_client() -> anthropic.Anthropic:
    """Return an Anthropic client, preferring ANTHROPIC_API_KEY env var."""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "my_api_key")
    return anthropic.Anthropic(api_key=api_key)


def _run_with_tools(client: anthropic.Anthropic, user_message: str, command: str) -> str:
    """
    Stream a response from Claude with web_search + local tool support.
    Returns the final text output.
    """
    tools = [
        {
            "type": "web_search_20250305",
            "name": "web_search",
            "user_location": {"type": "approximate", "region": "Middle East"},
        },
    ]

    messages = [{"role": "user", "content": user_message}]
    collected_text = []

    console.print(f"\n[cyan]AI processing: {command}...[/cyan]")

    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=128000,
        system=SYSTEM_PROMPT,
        tools=tools,
        messages=messages,
        thinking={"type": "adaptive"},
    ) as stream:
        for text in stream.text_stream:
            print(text, end="", flush=True)
            collected_text.append(text)

    print()
    return "".join(collected_text)


def run_snapshot(client: anthropic.Anthropic, url: str, client_name: str, magazine: str,
                 headline: str = "", date_str: str = "") -> None:
    """Command 1: AI-assisted snapshot with web search for metadata, then local doc generation."""
    console.print(
        Panel(
            f"[bold cyan]AI Snapshot Agent[/bold cyan]\n"
            f"URL      : {url}\n"
            f"Client   : [yellow]{client_name}[/yellow]\n"
            f"Magazine : [yellow]{magazine}[/yellow]",
            title="ADMC Media Monitor",
            border_style="blue",
        )
    )

    # Ask Claude to look up the article details via web search
    user_msg = (
        f"Look up this article and extract its headline, publication date, and publication name.\n"
        f"URL: {url}\n"
        f"Client: {client_name}\n"
        f"Magazine hint: {magazine}\n"
        f"Then confirm the details I should use for the snapshot report:\n"
        f"- Publication name\n- Headline\n- Date of publishing\n- Full URL\n\n"
        f"Provide a structured summary so the snapshot report can be created."
    )

    ai_summary = _run_with_tools(client, user_msg, "snapshot metadata lookup")

    # Now run the local snapshot tool with whatever we know
    console.print("\n[cyan]Generating Word document locally...[/cyan]")
    from media_monitor.snapshot import create_snapshot
    out_path = create_snapshot(
        url=url,
        client=client_name,
        magazine=magazine,
        headline=headline,
        date_str=date_str or datetime.today().strftime("%d %B %Y"),
        publication=magazine,
    )
    console.print(f"[bold green]Snapshot saved:[/bold green] {out_path}")


def run_coverage_24h(client: anthropic.Anthropic, client_name: str, campaign: str,
                     items_file: str | None = None, links: list[str] | None = None,
                     date_str: str = "") -> None:
    """Command 2: AI-assisted 24-hour coverage report."""
    console.print(
        Panel(
            f"[bold cyan]AI 24h Coverage Agent[/bold cyan]\n"
            f"Client   : [yellow]{client_name}[/yellow]\n"
            f"Campaign : [yellow]{campaign}[/yellow]",
            title="ADMC Media Monitor",
            border_style="blue",
        )
    )

    # Load items from file if provided
    items = []
    if items_file and os.path.exists(items_file):
        with open(items_file, "r", encoding="utf-8") as fh:
            items = json.load(fh)
        console.print(f"[green]Loaded {len(items)} items from {items_file}[/green]")

    # If links provided (or use KNOWBE4_LINKS default), ask Claude to analyze them
    source_links = links or (KNOWBE4_LINKS if not items else [])
    if source_links and not items:
        links_block = "\n".join(f"- {lnk}" for lnk in source_links)
        user_msg = (
            f"For client '{client_name}', campaign '{campaign}', analyze these coverage links "
            f"and extract structured data for each:\n{links_block}\n\n"
            f"For each article provide: publication name, headline, date (DD-MM-YYYY), "
            f"URL, language (English/Arabic), print_online (Online/Print).\n"
            f"Return a JSON array with those fields."
        )
        ai_response = _run_with_tools(client, user_msg, "coverage link analysis")

        # Try to parse JSON from AI response
        try:
            start = ai_response.find("[")
            end = ai_response.rfind("]") + 1
            if start != -1 and end > start:
                items = json.loads(ai_response[start:end])
                console.print(f"[green]AI extracted {len(items)} coverage items.[/green]")
        except json.JSONDecodeError:
            console.print("[yellow]Could not parse AI JSON — loading from coverage_items.json[/yellow]")
            fallback = os.path.join(os.path.dirname(__file__), "coverage_items.json")
            if os.path.exists(fallback):
                with open(fallback, "r", encoding="utf-8") as fh:
                    items = json.load(fh)

    if not items:
        console.print("[red]No coverage items found. Aborting.[/red]")
        return

    # Generate the report locally
    from media_monitor.coverage_24h import create_coverage_24h
    out_path = create_coverage_24h(
        client=client_name,
        campaign=campaign,
        items=items,
        date_str=date_str or datetime.today().strftime("%d %B %Y"),
    )
    console.print(f"[bold green]24h Coverage Report saved:[/bold green] {out_path}")


def run_tracker_update(client: anthropic.Anthropic, client_name: str,
                       entries_file: str | None = None) -> None:
    """Command 3: AI-assisted Excel tracker update."""
    console.print(
        Panel(
            f"[bold cyan]AI Tracker Agent[/bold cyan]\n"
            f"Client  : [yellow]{client_name}[/yellow]",
            title="ADMC Media Monitor",
            border_style="blue",
        )
    )

    entries = []
    if entries_file and os.path.exists(entries_file):
        with open(entries_file, "r", encoding="utf-8") as fh:
            entries = json.load(fh)
        console.print(f"[green]Loaded {len(entries)} entries from {entries_file}[/green]")
    else:
        # Fall back to bundled example
        fallback = os.path.join(os.path.dirname(__file__), "tracker_entries.json")
        if os.path.exists(fallback):
            with open(fallback, "r", encoding="utf-8") as fh:
                entries = json.load(fh)
            console.print(f"[green]Using fallback tracker_entries.json ({len(entries)} rows)[/green]")

    if not entries:
        console.print("[red]No tracker entries found. Aborting.[/red]")
        return

    from media_monitor.tracker import update_tracker
    out_path = update_tracker(client=client_name, entries=entries)
    console.print(f"[bold green]Tracker updated:[/bold green] {out_path}")


# ─── CLI ──────────────────────────────────────────────────────────────────────

@click.group()
def cli():
    """ADMC AI Media Monitor — Anthropic-powered PR reporting agent."""
    pass


@cli.command("snapshot")
@click.option("--url",      required=True,  help="Article URL")
@click.option("--client",   required=True,  help="Client name  e.g. KnowBe4")
@click.option("--magazine", required=True,  help="Publication name")
@click.option("--headline", default="",     help="Article headline")
@click.option("--date",     default="",     help="Publishing date")
def cmd_snapshot(url, client, magazine, headline, date):
    """Take an AI-assisted snapshot of a media article."""
    anthropic_client = _build_client()
    run_snapshot(anthropic_client, url, client, magazine, headline, date)


@cli.command("coverage-24h")
@click.option("--client",     required=True, help="Client name")
@click.option("--campaign",   required=True, help="Press release / campaign name")
@click.option("--date",       default="",    help="Report date (defaults to today)")
@click.option("--items-file", default=None,  help="JSON file with coverage items")
@click.option("--links",      default=None,  help="Comma-separated list of coverage URLs")
def cmd_coverage(client, campaign, date, items_file, links):
    """Build an AI-assisted 24-hour press coverage report."""
    anthropic_client = _build_client()
    link_list = [l.strip() for l in links.split(",")] if links else None
    run_coverage_24h(anthropic_client, client, campaign, items_file, link_list, date)


@cli.command("tracker")
@click.option("--client",       required=True, help="Client name")
@click.option("--entries-file", default=None,  help="JSON file with tracker entries")
def cmd_tracker(client, entries_file):
    """Add entries to the AI-powered Excel tracker."""
    anthropic_client = _build_client()
    run_tracker_update(anthropic_client, client, entries_file)


@cli.command("demo")
@click.option("--client",   default="KnowBe4",
              help="Client name (default: KnowBe4)")
@click.option("--campaign", default="Phish Alert Button for Microsoft Teams",
              help="Campaign/press-release name")
def cmd_demo(client, campaign):
    """
    Run the full demo: coverage-24h + tracker for the bundled KnowBe4 links.

    \b
    Example:
      python anthropic_client.py demo
      python anthropic_client.py demo --client "KnowBe4" --campaign "Phish Alert Button"
    """
    anthropic_client = _build_client()
    console.print(
        Panel(
            "[bold cyan]ADMC Media Monitor — Full Demo[/bold cyan]\n"
            "Running: coverage-24h + tracker update",
            title="ADMC",
            border_style="green",
        )
    )
    run_coverage_24h(anthropic_client, client, campaign, links=KNOWBE4_LINKS)
    run_tracker_update(anthropic_client, client)


if __name__ == "__main__":
    cli()
