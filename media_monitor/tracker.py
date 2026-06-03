"""
Command 3 - Excel Tracker
Adds or updates coverage entries in the ADMC Excel tracker.

Saved as: ADMC_clients_<ClientName>tracker.xlsx
"""

import os
from datetime import datetime

from openpyxl.utils import get_column_letter
from rich.console import Console
from rich.panel import Panel
from rich.table import Table as RichTable

from media_monitor.config import TRACKER_DIR, ACTIVE_LOGO_PATH, LOGOS_DIR
from media_monitor.utils.file_naming import tracker_filename
from media_monitor.utils.excel_helper import (
    create_or_load_tracker,
    append_tracker_row,
    next_sn,
)

console = Console()


def _insert_logos_excel(wb, ws, client_name: str) -> None:
    """
    Insert client logo and Active logo into the Excel sheet header area.
    Requires openpyxl + Pillow.
    """
    try:
        from openpyxl.drawing.image import Image as XLImage

        # Client logo
        safe_name = client_name.lower().replace(" ", "_")
        client_logo = None
        for ext in [".png", ".jpg", ".jpeg"]:
            candidate = os.path.join(LOGOS_DIR, f"{safe_name}{ext}")
            if os.path.exists(candidate):
                client_logo = candidate
                break

        if client_logo:
            img = XLImage(client_logo)
            img.width  = 130
            img.height = 50
            img.anchor = "A1"
            ws.add_image(img)

        # Active logo
        if os.path.exists(ACTIVE_LOGO_PATH):
            img2 = XLImage(ACTIVE_LOGO_PATH)
            img2.width  = 100
            img2.height = 40
            img2.anchor = "K1"
            ws.add_image(img2)

    except Exception as e:
        console.print(f"[yellow]Warning: Could not insert logos into Excel: {e}[/yellow]")


def update_tracker(client: str, entries: list[dict]) -> str:
    """
    Main entry point for Command 3.

    `entries` is a list of dicts matching TRACKER_COLUMNS keys.
    Returns the path to the updated .xlsx file.
    """
    console.print(
        Panel(
            f"[bold cyan]Updating Excel Tracker[/bold cyan]\n"
            f"Client  : [yellow]{client}[/yellow]\n"
            f"Entries : [yellow]{len(entries)}[/yellow]",
            title="ADMC Media Monitor",
            border_style="blue",
        )
    )

    filepath = tracker_filename(client)
    is_new   = not os.path.exists(filepath)

    wb, ws = create_or_load_tracker(filepath)

    # Insert logos on first creation
    if is_new:
        _insert_logos_excel(wb, ws, client)

    # Append entries
    added = 0
    for entry in entries:
        sn = next_sn(ws)
        append_tracker_row(ws, entry, sn)
        added += 1

    wb.save(filepath)

    # Rich summary table
    table = RichTable(title=f"Added {added} row(s) to tracker", show_lines=True)
    table.add_column("S/N",         style="cyan",   justify="center")
    table.add_column("Date",        style="white")
    table.add_column("Publication", style="yellow")
    table.add_column("Headline",    style="white",  no_wrap=False)
    table.add_column("Language",    style="green",  justify="center")

    for i, entry in enumerate(entries, start=1):
        headline = entry.get("Headline", "")
        table.add_row(
            str(i),
            entry.get("Date (DD/MM/YYYY)", ""),
            entry.get("Publication", ""),
            headline[:60] + ("..." if len(headline) > 60 else ""),
            entry.get("Language", ""),
        )

    console.print(table)
    console.print(
        f"\n[bold green]Tracker updated:[/bold green] "
        f"[underline]{filepath}[/underline]"
    )
    return filepath
