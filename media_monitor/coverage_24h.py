"""
Command 2 - 24-Hour Coverage Report
Aggregates all coverage items collected within the last 24 hours
for a given client / campaign and produces a formatted Word document.

Saved as: ADMC_<Client>_<PressReleaseName>_<Date>.docx
"""

import os
from datetime import datetime

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from rich.console import Console
from rich.panel import Panel

from media_monitor.config import COVERAGE_DIR, COVERAGE_COLUMNS
from media_monitor.utils.file_naming import coverage_filename
from media_monitor.utils.logo_handler import add_logo_header

console = Console()


def _shading_element(hex_color: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    return shd


def _add_title_block(doc: Document, client: str, campaign: str, total: int,
                     print_count: int, english_count: int, arabic_count: int) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(campaign)
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x00, 0x5B, 0xB5)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    total_run = p2.add_run(
        f"Total Secured Coverage {total}  "
        f"(Print {print_count}, English {english_count}, Arabic {arabic_count})"
    )
    total_run.bold      = True
    total_run.font.size = Pt(11)


def _add_section_heading(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x00, 0x5B, 0xB5)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(4)

    pPr  = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "00AEEF")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_snapshot_entry(doc: Document, sn: int, item: dict) -> None:
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_after = Pt(4)

    pub_run = p.add_run(f"{item.get('publication', '')}  -  {item.get('date', '')}")
    pub_run.bold      = True
    pub_run.font.size = Pt(10)

    p.add_run(" - ")

    hl_run = p.add_run(item.get("headline", ""))
    hl_run.font.color.rgb = RGBColor(0x00, 0x5B, 0xB5)
    hl_run.font.underline = True
    hl_run.font.size      = Pt(10)

    screenshot = item.get("screenshot_path")
    if screenshot and os.path.exists(screenshot):
        img_p = doc.add_paragraph()
        img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = img_p.add_run()
        run.add_picture(screenshot, width=Inches(5.8))
        doc.add_paragraph()


def _add_coverage_table(doc: Document, items: list[dict]) -> None:
    _add_section_heading(doc, "Full Coverage List")

    table = doc.add_table(rows=1, cols=len(COVERAGE_COLUMNS))
    table.style = "Table Grid"

    hdr = table.rows[0]
    for col_idx, col_name in enumerate(COVERAGE_COLUMNS):
        cell = hdr.cells[col_idx]
        cell._tc.get_or_add_tcPr().append(_shading_element("1F497D"))
        p    = cell.paragraphs[0]
        run  = p.add_run(col_name)
        run.bold            = True
        run.font.color.rgb  = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size       = Pt(9)
        p.alignment         = WD_ALIGN_PARAGRAPH.CENTER

    for sn, item in enumerate(items, start=1):
        row = table.add_row()
        values = [
            str(sn),
            item.get("publication", ""),
            item.get("headline", ""),
            item.get("language", "English"),
            item.get("date", ""),
            item.get("url", ""),
            item.get("print_online", "Online"),
        ]
        fill = "F2F2F2" if sn % 2 == 0 else "FFFFFF"
        for col_idx, val in enumerate(values):
            cell = row.cells[col_idx]
            cell._tc.get_or_add_tcPr().append(_shading_element(fill))
            p    = cell.paragraphs[0]
            run  = p.add_run(val)
            run.font.size = Pt(8)
            p.alignment   = WD_ALIGN_PARAGRAPH.CENTER if col_idx in [0, 3, 6] \
                            else WD_ALIGN_PARAGRAPH.LEFT


def create_coverage_24h(
    client: str,
    campaign: str,
    items: list[dict],
    date_str: str = "",
) -> str:
    console.print(
        Panel(
            f"[bold cyan]Creating 24h Coverage Report[/bold cyan]\n"
            f"Client   : [yellow]{client}[/yellow]\n"
            f"Campaign : [yellow]{campaign}[/yellow]\n"
            f"Items    : [yellow]{len(items)}[/yellow]",
            title="ADMC Media Monitor",
            border_style="blue",
        )
    )

    date_str = date_str or datetime.today().strftime("%d %B %Y")

    total         = len(items)
    print_count   = sum(1 for i in items if i.get("print_online", "").lower() == "print")
    english_count = sum(1 for i in items if i.get("language", "English").lower() == "english")
    arabic_count  = sum(1 for i in items if i.get("language", "").lower() == "arabic")

    print_items   = [i for i in items if i.get("print_online", "").lower() == "print"]
    english_items = [i for i in items if i.get("language", "English").lower() == "english"
                     and i.get("print_online", "Online").lower() != "print"]
    arabic_items  = [i for i in items if i.get("language", "").lower() == "arabic"
                     and i.get("print_online", "Online").lower() != "print"]

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin   = Inches(0.7)
        section.right_margin  = Inches(0.7)

    add_logo_header(doc, client)

    _add_title_block(doc, client, campaign, total,
                     print_count, english_count, arabic_count)
    doc.add_paragraph()

    if print_items:
        _add_section_heading(doc, "Print:")
        for sn, item in enumerate(print_items, start=1):
            _add_snapshot_entry(doc, sn, item)

    if english_items:
        _add_section_heading(doc, "English Online:")
        for sn, item in enumerate(english_items, start=1):
            _add_snapshot_entry(doc, sn, item)

    if arabic_items:
        _add_section_heading(doc, "Arabic Online:")
        for sn, item in enumerate(arabic_items, start=1):
            _add_snapshot_entry(doc, sn, item)

    _add_coverage_table(doc, items)

    out_path = coverage_filename(client, campaign, date_str)
    doc.save(out_path)

    console.print(
        f"\n[bold green]24h Coverage Report created:[/bold green] "
        f"[underline]{out_path}[/underline]"
    )
    return out_path
