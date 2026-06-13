#!/usr/bin/env python3
# /// script
# requires-python = ">=3.9"
# dependencies = ["python-docx"]
# ///
"""Doc Manager v2.0 — Professional AI Document Generation Tool.

Create and edit polished, business-ready Microsoft Word (.docx) documents
with themes, cover pages, table of contents, callout blocks, enhanced tables,
headers/footers, and more.

Usage:
    uv run scripts/docx_tool.py create --title "Report" --content content.json --theme professional
    uv run scripts/docx_tool.py edit --file doc.docx --action append --content extra.json
    uv run scripts/docx_tool.py edit --file doc.docx --action replace --find "old" --replace-with "new"
    uv run scripts/docx_tool.py edit --file doc.docx --action insert-after --heading "Intro" --content new.json
    uv run scripts/docx_tool.py edit --file doc.docx --action replace-section --heading "Intro" --content new.json
    uv run scripts/docx_tool.py edit --file doc.docx --action delete-section --heading "Appendix"
    uv run scripts/docx_tool.py add-table --file doc.docx --table-data table.json
    uv run scripts/docx_tool.py remove-table --file doc.docx --index 0
    uv run scripts/docx_tool.py info --file doc.docx --output info.json
    uv run scripts/docx_tool.py validate --content content.json
"""

import argparse
import datetime
import json
import os
import sys
import subprocess
import shutil
import csv

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ============================================================================
# SECTION 1: CONSTANTS & FONT DETECTION
# ============================================================================

DEFAULT_OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Downloads")

SUPPORTED_BLOCK_TYPES = {
    "heading", "paragraph", "bullet_list", "numbered_list", "table",
    "cover_page", "toc", "executive_summary",
    "note", "warning", "important",
    "code", "checklist", "timeline", "page_break",
    "image", "chart",
}


_font_cache = {}

def _detect_font(preferences=None):
    """Detect the best available font from preference list via Windows Registry."""
    if preferences is None:
        preferences = ["Aptos", "Calibri", "Arial"]
    
    cache_key = tuple(preferences)
    if cache_key in _font_cache:
        return _font_cache[cache_key]

    result = preferences[1] if len(preferences) > 1 else "Calibri"
    if sys.platform == "win32":
        try:
            import winreg
            key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                num_values = winreg.QueryInfoKey(key)[1]
                installed = set()
                for i in range(num_values):
                    name, _, _ = winreg.EnumValue(key, i)
                    installed.add(name.lower())
                for pref in preferences:
                    for name in installed:
                        if pref.lower() in name:
                            result = pref
                            break
                    else:
                        continue
                    break
        except Exception:
            pass
            
    _font_cache[cache_key] = result
    return result

def _replace_in_paragraph(paragraph, find_text, replace_text):
    """Replace text across multiple runs in a paragraph."""
    full = paragraph.text
    if find_text not in full:
        return 0
    new_text = full.replace(find_text, replace_text)
    if paragraph.runs:
        for i, run in enumerate(paragraph.runs):
            if i == 0:
                run.text = new_text
            else:
                run.text = ""
    return full.count(find_text)


# ============================================================================
# SECTION 2: THEME SYSTEM
# ============================================================================

THEMES = {
    "professional": {
        "name": "Professional",
        "body_font": None,  # resolved dynamically
        "body_size": 11,
        "code_font": "Consolas",
        "code_size": 10,
        "line_spacing": 1.15,
        "para_spacing_after": 6,
        "heading_1": {"size": 24, "bold": True, "color": "1F4E79"},
        "heading_2": {"size": 18, "bold": True, "color": "333333"},
        "heading_3": {"size": 14, "bold": True, "color": "333333"},
        "heading_4": {"size": 12, "bold": True, "color": "555555"},
        "text_color": "333333",
        "accent_primary": "1F4E79",
        "accent_secondary": "2F75B5",
        "table_header_bg": "1F4E79",
        "table_header_text": "FFFFFF",
        "table_alt_row": "F2F7FB",
        "note_bg": "E8F4FD",
        "note_border": "2F75B5",
        "warning_bg": "FFF3CD",
        "warning_border": "FFC107",
        "important_bg": "F8D7DA",
        "important_border": "DC3545",
        "exec_summary_bg": "E8F4FD",
        "exec_summary_border": "1F4E79",
        "code_bg": "F5F5F5",
        "code_border": "E0E0E0",
        "cover_title_size": 36,
        "cover_subtitle_size": 18,
    },
    "technical": {
        "name": "Technical",
        "body_font": "Segoe UI",
        "body_size": 11,
        "code_font": "Consolas",
        "code_size": 10,
        "line_spacing": 1.15,
        "para_spacing_after": 6,
        "heading_1": {"size": 22, "bold": True, "color": "2D2D2D"},
        "heading_2": {"size": 17, "bold": True, "color": "2D2D2D"},
        "heading_3": {"size": 13, "bold": True, "color": "444444"},
        "heading_4": {"size": 11, "bold": True, "color": "444444"},
        "text_color": "333333",
        "accent_primary": "0078D4",
        "accent_secondary": "50A0E0",
        "table_header_bg": "2D2D2D",
        "table_header_text": "FFFFFF",
        "table_alt_row": "F5F5F5",
        "note_bg": "E3F2FD",
        "note_border": "0078D4",
        "warning_bg": "FFF8E1",
        "warning_border": "FF8F00",
        "important_bg": "FCE4EC",
        "important_border": "D32F2F",
        "exec_summary_bg": "E3F2FD",
        "exec_summary_border": "0078D4",
        "code_bg": "1E1E1E",
        "code_border": "333333",
        "cover_title_size": 32,
        "cover_subtitle_size": 16,
    },
    "executive": {
        "name": "Executive",
        "body_font": None,
        "body_size": 11,
        "code_font": "Consolas",
        "code_size": 10,
        "line_spacing": 1.3,
        "para_spacing_after": 8,
        "heading_1": {"size": 28, "bold": True, "color": "1A1A2E"},
        "heading_2": {"size": 20, "bold": True, "color": "1A1A2E"},
        "heading_3": {"size": 15, "bold": True, "color": "333333"},
        "heading_4": {"size": 12, "bold": True, "color": "333333"},
        "text_color": "333333",
        "accent_primary": "1A1A2E",
        "accent_secondary": "16213E",
        "table_header_bg": "1A1A2E",
        "table_header_text": "FFFFFF",
        "table_alt_row": "F8F8FA",
        "note_bg": "EEF0F7",
        "note_border": "1A1A2E",
        "warning_bg": "FFF3CD",
        "warning_border": "FFC107",
        "important_bg": "F8D7DA",
        "important_border": "DC3545",
        "exec_summary_bg": "EEF0F7",
        "exec_summary_border": "1A1A2E",
        "code_bg": "F5F5F5",
        "code_border": "E0E0E0",
        "cover_title_size": 40,
        "cover_subtitle_size": 20,
    },
    "startup": {
        "name": "Startup",
        "body_font": None,
        "body_size": 11,
        "code_font": "Consolas",
        "code_size": 10,
        "line_spacing": 1.2,
        "para_spacing_after": 6,
        "heading_1": {"size": 26, "bold": True, "color": "6C5CE7"},
        "heading_2": {"size": 19, "bold": True, "color": "2D3436"},
        "heading_3": {"size": 14, "bold": True, "color": "636E72"},
        "heading_4": {"size": 12, "bold": True, "color": "636E72"},
        "text_color": "2D3436",
        "accent_primary": "6C5CE7",
        "accent_secondary": "A29BFE",
        "table_header_bg": "6C5CE7",
        "table_header_text": "FFFFFF",
        "table_alt_row": "F6F5FF",
        "note_bg": "EDE7F6",
        "note_border": "6C5CE7",
        "warning_bg": "FFF3CD",
        "warning_border": "FDCB6E",
        "important_bg": "FFEEF0",
        "important_border": "FF7675",
        "exec_summary_bg": "EDE7F6",
        "exec_summary_border": "6C5CE7",
        "code_bg": "F5F5F5",
        "code_border": "E0E0E0",
        "cover_title_size": 34,
        "cover_subtitle_size": 17,
    },
}


_resolved_themes = {}

def _resolve_theme(theme_name=None, content_data=None):
    """Resolve theme by name, falling back to content JSON or default."""
    name = theme_name
    if not name and isinstance(content_data, dict):
        name = content_data.get("theme")
    if not name:
        name = "professional"
    if name not in THEMES:
        print(f"Error: Unknown theme '{name}'. Available: {', '.join(THEMES.keys())}",
              file=sys.stderr)
        sys.exit(1)
        
    if name in _resolved_themes:
        return _resolved_themes[name]
        
    theme = dict(THEMES[name])
    if theme["body_font"] is None:
        theme["body_font"] = _detect_font()
        
    _resolved_themes[name] = theme
    return theme


# ============================================================================
# SECTION 3: UTILITY HELPERS
# ============================================================================

def _hex_to_rgb(hex_str):
    """Convert hex color string (e.g. '1F4E79') to RGBColor."""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_run_font(run, font_name, font_size_pt, bold=False, italic=False, color=None):
    """Apply font styling to a single run."""
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = _hex_to_rgb(color)


def _set_style_font_xml(style_element, font_name):
    """Apply XML-level font references for cross-platform compatibility on styles."""
    rPr = style_element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)



# ============================================================================
# SECTION 4: XML HELPERS
# ============================================================================

def _set_paragraph_shading(paragraph, hex_color):
    """Apply background shading to an entire paragraph."""
    pPr = paragraph._element.get_or_add_pPr()
    existing = pPr.find(qn("w:shd"))
    if existing is not None:
        pPr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def _set_paragraph_borders(paragraph, sides=None, color="4472C4", size="24", space="10"):
    """Apply borders to a paragraph. size is in 1/8pt units (24 = 3pt)."""
    if sides is None:
        sides = ["left"]
    pPr = paragraph._element.get_or_add_pPr()
    existing = pPr.find(qn("w:pBdr"))
    if existing is not None:
        pPr.remove(existing)
    pBdr = OxmlElement("w:pBdr")
    for side in sides:
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), space)
        border.set(qn("w:color"), color)
        pBdr.append(border)
    pPr.append(pBdr)


def _set_cell_shading(cell, hex_color):
    """Apply background shading to a table cell. MUST create new element per cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:shd"))
    if existing is not None:
        tcPr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_cell_borders(cell, accent_color, left_size="24", other_color="E0E0E0", other_size="4"):
    """Set cell borders: thick colored left + thin gray others."""
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:tcBorders"))
    if existing is not None:
        tcPr.remove(existing)
    borders = OxmlElement("w:tcBorders")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), left_size)
    left.set(qn("w:space"), "0")
    left.set(qn("w:color"), accent_color)
    borders.append(left)
    for side in ("top", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), other_size)
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), other_color)
        borders.append(b)
    tcPr.append(borders)


def _set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    """Set internal cell margins in twips."""
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:tcMar"))
    if existing is not None:
        tcPr.remove(existing)
    margins = OxmlElement("w:tcMar")
    for name, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        m = OxmlElement(f"w:{name}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        margins.append(m)
    tcPr.append(margins)


def _set_table_full_width(table):
    """Set table to span the full page width (100%)."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    existing = tblPr.find(qn("w:tblW"))
    if existing is not None:
        tblPr.remove(existing)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "5000")
    tblW.set(qn("w:type"), "pct")
    tblPr.append(tblW)


def _remove_table_borders(table):
    """Remove all borders from a table (used for callout boxes)."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    existing = tblPr.find(qn("w:tblBorders"))
    if existing is not None:
        tblPr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        borders.append(b)
    tblPr.append(borders)


def _add_field_code(paragraph, field_type, theme=None, font_size=None, color=None):
    """Insert a Word field code (PAGE, NUMPAGES, TOC, etc.) into a paragraph."""
    run1 = paragraph.add_run()
    fc_begin = OxmlElement("w:fldChar")
    fc_begin.set(qn("w:fldCharType"), "begin")
    run1._r.append(fc_begin)

    run2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_type} "
    run2._r.append(instr)

    run3 = paragraph.add_run()
    fc_sep = OxmlElement("w:fldChar")
    fc_sep.set(qn("w:fldCharType"), "separate")
    run3._r.append(fc_sep)

    run4 = paragraph.add_run()
    fc_end = OxmlElement("w:fldChar")
    fc_end.set(qn("w:fldCharType"), "end")
    run4._r.append(fc_end)

    if theme and font_size and color:
        for r in (run1, run2, run3, run4):
            _set_run_font(r, theme["body_font"], font_size, color=color)


def _set_tab_stops(paragraph, center_pos=4680, right_pos=9360):
    """Set center and right tab stops on a paragraph (positions in twips)."""
    pPr = paragraph._element.get_or_add_pPr()
    existing = pPr.find(qn("w:tabs"))
    if existing is not None:
        pPr.remove(existing)
    tabs = OxmlElement("w:tabs")
    tc = OxmlElement("w:tab")
    tc.set(qn("w:val"), "center")
    tc.set(qn("w:pos"), str(center_pos))
    tabs.append(tc)
    tr = OxmlElement("w:tab")
    tr.set(qn("w:val"), "right")
    tr.set(qn("w:pos"), str(right_pos))
    tabs.append(tr)
    pPr.append(tabs)


# ============================================================================
# SECTION 5: CONTENT BLOCK PROCESSORS
# ============================================================================

def _add_heading(doc, theme, text, level=1):
    """Add a themed heading paragraph."""
    heading = doc.add_heading("", level=level)
    heading.clear()
    run = heading.add_run(text)
    h_conf = theme.get(f"heading_{level}", {"size": 12, "bold": True, "color": "333333"})
    _set_run_font(run, theme["body_font"], h_conf["size"],
                  bold=h_conf.get("bold", True), color=h_conf["color"])
    return heading


def _add_paragraph_block(doc, theme, text, bold=False, italic=False):
    """Add a themed body paragraph."""
    para = doc.add_paragraph()
    run = para.add_run(text)
    _set_run_font(run, theme["body_font"], theme["body_size"],
                  bold=bold, italic=italic, color=theme["text_color"])
    return para


def _add_bullet_list(doc, theme, items):
    """Add an unordered bullet list."""
    for item in items:
        para = doc.add_paragraph(style="List Bullet")
        para.clear()
        run = para.add_run(item)
        _set_run_font(run, theme["body_font"], theme["body_size"], color=theme["text_color"])


def _add_numbered_list(doc, theme, items):
    """Add an ordered numbered list."""
    for item in items:
        para = doc.add_paragraph(style="List Number")
        para.clear()
        run = para.add_run(item)
        _set_run_font(run, theme["body_font"], theme["body_size"], color=theme["text_color"])


def _add_table_block(doc, theme, headers, rows):
    """Add a professionally styled table with colored header and alternating rows."""
    if not headers:
        print("Warning: Table block has empty headers, skipping.", file=sys.stderr)
        return
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    _set_table_full_width(table)

    # Header row
    for i, header_text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(header_text)
        _set_run_font(run, theme["body_font"], theme["body_size"],
                      bold=True, color=theme["table_header_text"])
        _set_cell_shading(cell, theme["table_header_bg"])

    # Data rows with alternating shading
    for row_idx, row_data in enumerate(rows):
        row = table.rows[row_idx + 1]
        is_alt = row_idx % 2 == 1
        for col_idx, cell_text in enumerate(row_data):
            if col_idx < len(row.cells):
                cell = row.cells[col_idx]
                cell.text = ""
                run = cell.paragraphs[0].add_run(str(cell_text))
                _set_run_font(run, theme["body_font"], theme["body_size"],
                              color=theme["text_color"])
                if is_alt:
                    _set_cell_shading(cell, theme["table_alt_row"])
    return table


def _add_cover_page(doc, theme, block):
    """Add a centered cover page with title, subtitle, author, date."""
    title = block.get("title", "Untitled Document")
    subtitle = block.get("subtitle")
    author = block.get("author")
    date_str = block.get("date")

    # Vertical spacer
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(120)
    spacer.paragraph_format.space_after = Pt(0)

    # Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(title)
    _set_run_font(title_run, theme["body_font"], theme.get("cover_title_size", 36),
                  bold=True, color=theme["accent_primary"])
    title_para.paragraph_format.space_after = Pt(8)

    # Subtitle
    if subtitle:
        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_run = sub_para.add_run(subtitle)
        _set_run_font(sub_run, theme["body_font"], theme.get("cover_subtitle_size", 18),
                      color=theme["accent_secondary"])
        sub_para.paragraph_format.space_after = Pt(20)

    # Separator line
    sep_para = doc.add_paragraph()
    sep_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sep_run = sep_para.add_run("\u2501" * 30)  # ━ heavy horizontal line
    _set_run_font(sep_run, theme["body_font"], 12, color=theme["accent_secondary"])
    sep_para.paragraph_format.space_before = Pt(20)
    sep_para.paragraph_format.space_after = Pt(20)

    # Author
    if author:
        auth_para = doc.add_paragraph()
        auth_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        auth_run = auth_para.add_run(author)
        _set_run_font(auth_run, theme["body_font"], 14, color=theme["text_color"])
        auth_para.paragraph_format.space_after = Pt(4)

    # Date
    if date_str:
        date_para = doc.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_run = date_para.add_run(date_str)
        _set_run_font(date_run, theme["body_font"], 12, color="888888")


def _add_toc(doc, theme):
    """Add a Table of Contents field code. Word must update it (Ctrl+A, F9)."""
    _add_heading(doc, theme, "Table of Contents", level=1)
    para = doc.add_paragraph()

    # TOC field: begin
    run1 = para.add_run()
    fc1 = OxmlElement("w:fldChar")
    fc1.set(qn("w:fldCharType"), "begin")
    run1._r.append(fc1)

    # TOC field: instruction
    run2 = para.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    run2._r.append(instr)

    # TOC field: separate
    run3 = para.add_run()
    fc2 = OxmlElement("w:fldChar")
    fc2.set(qn("w:fldCharType"), "separate")
    run3._r.append(fc2)

    # Placeholder text (shown until user updates the field)
    run4 = para.add_run("Update this field in Word to see Table of Contents (Ctrl+A, F9)")
    _set_run_font(run4, theme["body_font"], theme["body_size"], italic=True, color="999999")

    # TOC field: end
    run5 = para.add_run()
    fc3 = OxmlElement("w:fldChar")
    fc3.set(qn("w:fldCharType"), "end")
    run5._r.append(fc3)

    # Page break after TOC
    doc.add_page_break()


def _add_executive_summary(doc, theme, text):
    """Add an executive summary callout box."""
    table = doc.add_table(rows=1, cols=1)
    _set_table_full_width(table)
    _remove_table_borders(table)
    cell = table.rows[0].cells[0]
    cell.paragraphs[0].clear()

    # Title
    title_run = cell.paragraphs[0].add_run("Executive Summary")
    _set_run_font(title_run, theme["body_font"], 14, bold=True, color=theme["accent_primary"])
    cell.paragraphs[0].paragraph_format.space_after = Pt(6)

    # Content
    content_para = cell.add_paragraph()
    content_run = content_para.add_run(text)
    _set_run_font(content_run, theme["body_font"], theme["body_size"], color=theme["text_color"])
    content_para.paragraph_format.space_after = Pt(4)

    # Style the cell
    _set_cell_shading(cell, theme["exec_summary_bg"])
    _set_cell_borders(cell, theme["exec_summary_border"])
    _set_cell_margins(cell, top=100, bottom=100, left=160, right=120)


def _add_callout(doc, theme, callout_type, title, text):
    """Add a callout block (note, warning, or important)."""
    colors_map = {
        "note": ("note_bg", "note_border", "\u2139\ufe0f"),       # ℹ️
        "warning": ("warning_bg", "warning_border", "\u26a0\ufe0f"),  # ⚠️
        "important": ("important_bg", "important_border", "\u2757"),   # ❗
    }
    bg_key, border_key, icon = colors_map.get(callout_type, colors_map["note"])
    bg_color = theme[bg_key]
    border_color = theme[border_key]
    display_title = title or callout_type.capitalize()

    table = doc.add_table(rows=1, cols=1)
    _set_table_full_width(table)
    _remove_table_borders(table)
    cell = table.rows[0].cells[0]
    cell.paragraphs[0].clear()

    # Title line
    title_run = cell.paragraphs[0].add_run(f"{icon}  {display_title}")
    _set_run_font(title_run, theme["body_font"], theme["body_size"], bold=True,
                  color=theme["text_color"])
    cell.paragraphs[0].paragraph_format.space_after = Pt(4)

    # Content
    content_para = cell.add_paragraph()
    content_run = content_para.add_run(text)
    _set_run_font(content_run, theme["body_font"], theme["body_size"], color=theme["text_color"])
    content_para.paragraph_format.space_after = Pt(4)

    _set_cell_shading(cell, bg_color)
    _set_cell_borders(cell, border_color)
    _set_cell_margins(cell, top=80, bottom=80, left=160, right=120)


def _add_code_block(doc, theme, text, language=None):
    """Add a code block with monospace font and shaded background."""
    # Optional language label
    if language:
        label_para = doc.add_paragraph()
        label_run = label_para.add_run(f"  {language}")
        _set_run_font(label_run, theme["body_font"], 8, bold=True, color="666666")
        label_para.paragraph_format.space_after = Pt(0)
        label_para.paragraph_format.space_before = Pt(6)

    # Code content — each line as part of one paragraph
    code_para = doc.add_paragraph()
    code_run = code_para.add_run(text)
    _set_run_font(code_run, theme["code_font"], theme["code_size"], color=theme["text_color"])
    _set_paragraph_shading(code_para, theme["code_bg"])

    # Add left border and indentation for visual distinction
    _set_paragraph_borders(code_para, sides=["left", "top", "bottom", "right"],
                           color=theme["code_border"], size="4", space="6")

    # Preserve whitespace formatting
    code_para.paragraph_format.space_before = Pt(2) if not language else Pt(0)
    code_para.paragraph_format.space_after = Pt(6)

    # Set paragraph indentation for visual padding
    code_para.paragraph_format.left_indent = Pt(8)
    code_para.paragraph_format.right_indent = Pt(8)


def _add_checklist(doc, theme, items):
    """Add a checklist with Unicode checkboxes."""
    for item in items:
        para = doc.add_paragraph()
        if item.strip().lower().startswith("[x]"):
            checkbox = "\u2611 "  # ☑
            item_text = item.strip()[3:].strip()
        else:
            checkbox = "\u2610 "  # ☐
            raw = item.strip()
            if raw.startswith("[ ]"):
                item_text = raw[3:].strip()
            else:
                item_text = raw

        check_run = para.add_run(checkbox)
        _set_run_font(check_run, theme["body_font"], theme["body_size"], color=theme["accent_primary"])
        text_run = para.add_run(item_text)
        _set_run_font(text_run, theme["body_font"], theme["body_size"], color=theme["text_color"])
        para.paragraph_format.space_after = Pt(2)


def _add_timeline(doc, theme, items):
    """Add a timeline as a professionally styled milestone table."""
    # Determine columns based on whether description is present
    has_desc = any(item.get("description") for item in items)
    if has_desc:
        headers = ["Phase", "Date", "Description"]
    else:
        headers = ["Phase", "Date"]

    rows = []
    for item in items:
        row = [item.get("phase", ""), item.get("date", "")]
        if has_desc:
            row.append(item.get("description", ""))
        rows.append(row)

    _add_table_block(doc, theme, headers, rows)


def _add_page_break(doc):
    """Add a simple page break."""
    doc.add_page_break()


def _add_image_block(doc, theme, block):
    """Add an image with optional caption and alignment."""
    path = block.get("path")
    if not path or not os.path.isfile(path):
        print(f"Warning: Image not found at '{path}', skipping.", file=sys.stderr)
        return

    width = block.get("width")
    
    try:
        if width:
            picture = doc.add_picture(path, width=Inches(float(width)))
        else:
            picture = doc.add_picture(path)
            
        # Alignment
        alignment_str = block.get("alignment", "center").lower()
        if alignment_str == "center":
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif alignment_str == "right":
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
        # Caption
        caption_text = block.get("caption")
        if caption_text:
            caption = doc.add_paragraph(caption_text)
            if alignment_str == "center":
                caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif alignment_str == "right":
                caption.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for run in caption.runs:
                _set_run_font(run, theme["body_font"], theme["body_size"] - 1, italic=True, color="666666")
    except Exception as e:
        print(f"Warning: Failed to add image '{path}' - {e}", file=sys.stderr)


def _add_chart_block(doc, theme, block):
    """Add a matplotlib chart as an image."""
    try:
        import matplotlib.pyplot as plt
        import io
    except ImportError:
        print("Warning: matplotlib not installed. Cannot render charts. Install with: pip install matplotlib", file=sys.stderr)
        return

    chart_type = block.get("chart_type", "bar")
    title = block.get("title", "")
    data = block.get("data", {})
    labels = data.get("labels", [])
    
    fig, ax = plt.subplots(figsize=(6, 4))
    
    # Theme colors mapping
    accent_primary = "#" + theme.get("accent_primary", "1F4E79")
    accent_secondary = "#" + theme.get("accent_secondary", "2F75B5")
    colors = [accent_primary, accent_secondary, "#5B9BD5", "#9DC3E6"]
    
    if chart_type == "bar":
        values = data.get("values", [])
        if values:
            ax.bar(labels, values, color=accent_primary)
        elif "datasets" in data:
            import numpy as np
            datasets = data["datasets"]
            x = np.arange(len(labels))
            width = 0.8 / len(datasets)
            for i, ds in enumerate(datasets):
                ax.bar(x + i*width - width*(len(datasets)-1)/2, ds.get("values", []), width, label=ds.get("label", ""), color=colors[i % len(colors)])
            ax.legend()
    elif chart_type == "pie":
        values = data.get("values", [])
        if values:
            ax.pie(values, labels=labels, autopct='%1.1f%%', colors=colors)
    elif chart_type == "line":
        if "datasets" in data:
            datasets = data["datasets"]
            for i, ds in enumerate(datasets):
                ax.plot(labels, ds.get("values", []), marker='o', label=ds.get("label", ""), color=colors[i % len(colors)])
            ax.legend()
        else:
            values = data.get("values", [])
            ax.plot(labels, values, marker='o', color=accent_primary)
            
    if title:
        ax.set_title(title, color="#" + theme.get("text_color", "333333"))
        
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=200, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)
    
    width = block.get("width")
    try:
        if width:
            doc.add_picture(buf, width=Inches(float(width)))
        else:
            doc.add_picture(buf)
            
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        caption_text = block.get("caption")
        if caption_text:
            caption = doc.add_paragraph(caption_text)
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in caption.runs:
                _set_run_font(run, theme["body_font"], theme["body_size"] - 1, italic=True, color="666666")
    except Exception as e:
        print(f"Warning: Failed to add chart - {e}", file=sys.stderr)


# ============================================================================
# SECTION 6: LAYOUT — STYLES, METADATA, HEADER, FOOTER
# ============================================================================

def _apply_default_styles(doc, theme):
    """Configure document-wide default styles based on theme."""
    style = doc.styles["Normal"]
    font = style.font
    font.name = theme["body_font"]
    font.size = Pt(theme["body_size"])
    font.color.rgb = _hex_to_rgb(theme["text_color"])

    # XML-level font references
    _set_style_font_xml(style.element, theme["body_font"])

    # Paragraph format
    pf = style.paragraph_format
    pf.space_after = Pt(theme["para_spacing_after"])
    pf.line_spacing = theme["line_spacing"]

    # Configure heading styles (font name and size; color is set per-run)
    for level in range(1, 5):
        h_key = f"heading_{level}"
        if h_key in theme:
            h_conf = theme[h_key]
            try:
                h_style = doc.styles[f"Heading {level}"]
                h_style.font.name = theme["body_font"]
                h_style.font.size = Pt(h_conf["size"])
                h_style.font.bold = h_conf.get("bold", True)
                # Set XML font refs for heading styles
                _set_style_font_xml(h_style.element, theme["body_font"])
            except KeyError:
                pass


def _apply_metadata(doc, metadata):
    """Set document core properties from metadata dict."""
    props = doc.core_properties
    if metadata.get("author"):
        props.author = metadata["author"]
    if metadata.get("title"):
        props.title = metadata["title"]
    if metadata.get("version"):
        props.version = metadata["version"]
    if metadata.get("status"):
        props.content_status = metadata["status"]
    if metadata.get("department"):
        props.category = metadata["department"]
    if metadata.get("subject"):
        props.subject = metadata["subject"]
    if metadata.get("keywords"):
        props.keywords = metadata["keywords"]
    props.modified = datetime.datetime.now()


def _apply_header(doc, section_idx, theme, header_config):
    """Apply header with document title and bottom separator."""
    section = doc.sections[section_idx]
    if section_idx > 0:
        section.header.is_linked_to_previous = False

    header = section.header
    para = header.paragraphs[0]
    para.clear()

    text = header_config.get("text", "")
    if text:
        run = para.add_run(text)
        _set_run_font(run, theme["body_font"], 9, color="999999")

    # Bottom border separator
    _set_paragraph_borders(para, sides=["bottom"], color="CCCCCC", size="4", space="4")
    para.paragraph_format.space_after = Pt(0)


def _apply_footer(doc, section_idx, theme, footer_config):
    """Apply footer with version, page numbers, and optional confidential label."""
    section = doc.sections[section_idx]
    if section_idx > 0:
        section.footer.is_linked_to_previous = False

    footer = section.footer
    para = footer.paragraphs[0]
    para.clear()

    # Top border separator
    _set_paragraph_borders(para, sides=["top"], color="CCCCCC", size="4", space="4")

    # Tab stops for three-column layout
    _set_tab_stops(para)

    font_size = 8
    font_color = "999999"

    # Left: Version
    version = footer_config.get("version")
    if version:
        run = para.add_run(f"Version {version}")
        _set_run_font(run, theme["body_font"], font_size, color=font_color)

    # Tab to center
    para.add_run("\t")

    # Center: Page numbers
    if footer_config.get("page_numbers", True):
        run = para.add_run("Page ")
        _set_run_font(run, theme["body_font"], font_size, color=font_color)
        _add_field_code(para, "PAGE", theme, font_size, font_color)
        run = para.add_run(" of ")
        _set_run_font(run, theme["body_font"], font_size, color=font_color)
        _add_field_code(para, "NUMPAGES", theme, font_size, font_color)

    # Tab to right
    para.add_run("\t")

    # Right: Confidential label
    if footer_config.get("confidential"):
        run = para.add_run("CONFIDENTIAL")
        _set_run_font(run, theme["body_font"], font_size, bold=True, color="CC0000")

    para.paragraph_format.space_before = Pt(4)


# ============================================================================
# SECTION 7: DOCUMENT BUILDER
# ============================================================================

def _substitute_variables(data, variables):
    """Recursively substitute {{key}} placeholders in all strings."""
    if isinstance(data, str):
        for key, value in variables.items():
            data = data.replace(f"{{{{{key}}}}}", str(value))
        return data
    elif isinstance(data, list):
        return [_substitute_variables(item, variables) for item in data]
    elif isinstance(data, dict):
        return {k: _substitute_variables(v, variables) for k, v in data.items()}
    return data


def _normalize_content(content_data):
    """Normalize content data to v2 format (backward compatible with v1 arrays)."""
    if isinstance(content_data, list):
        return {"content": content_data}
    return content_data


def _process_content_blocks(doc, theme, blocks):
    """Process an array of content block dicts and add them to the document."""
    for block in blocks:
        block_type = block.get("type")

        if block_type == "cover_page":
            # cover_page in the middle of content is treated as centered text + page break
            _add_cover_page(doc, theme, block)
            doc.add_page_break()
        elif block_type == "heading":
            _add_heading(doc, theme, block.get("text", ""), block.get("level", 1))
        elif block_type == "paragraph":
            _add_paragraph_block(doc, theme, block.get("text", ""),
                                 bold=block.get("bold", False),
                                 italic=block.get("italic", False))
        elif block_type == "bullet_list":
            _add_bullet_list(doc, theme, block.get("items", []))
        elif block_type == "numbered_list":
            _add_numbered_list(doc, theme, block.get("items", []))
        elif block_type == "table":
            _add_table_block(doc, theme, block.get("headers", []), block.get("rows", []))
        elif block_type == "toc":
            _add_toc(doc, theme)
        elif block_type == "executive_summary":
            _add_executive_summary(doc, theme, block.get("text", ""))
        elif block_type in ("note", "warning", "important"):
            _add_callout(doc, theme, block_type, block.get("title"), block.get("text", ""))
        elif block_type == "code":
            _add_code_block(doc, theme, block.get("text", ""), block.get("language"))
        elif block_type == "checklist":
            _add_checklist(doc, theme, block.get("items", []))
        elif block_type == "timeline":
            _add_timeline(doc, theme, block.get("items", []))
        elif block_type == "page_break":
            _add_page_break(doc)
        elif block_type == "image":
            _add_image_block(doc, theme, block)
        elif block_type == "chart":
            _add_chart_block(doc, theme, block)
        else:
            print(f"Warning: Unknown block type '{block_type}', skipping.", file=sys.stderr)


def _build_document(content_data, theme_override=None):
    """Orchestrate full document creation from content JSON."""
    content_json = _normalize_content(content_data)
    theme = _resolve_theme(theme_override, content_json)

    doc = Document()
    _apply_default_styles(doc, theme)

    blocks = content_json.get("content", [])
    has_cover = blocks and blocks[0].get("type") == "cover_page"

    if has_cover:
        # Process cover page into section 0
        _add_cover_page(doc, theme, blocks[0])
        # Section break — cover page is now section 0, content starts in section 1
        doc.add_section()
        # Vertical center alignment for cover page section
        sectPr = doc.sections[0]._sectPr
        vAlign = OxmlElement("w:vAlign")
        vAlign.set(qn("w:val"), "center")
        sectPr.append(vAlign)
        # Process remaining blocks
        _process_content_blocks(doc, theme, blocks[1:])
    else:
        _process_content_blocks(doc, theme, blocks)

    # Apply metadata
    metadata = content_json.get("metadata")
    if metadata:
        _apply_metadata(doc, metadata)
        # Also set document title from metadata if not set
        if metadata.get("title") and not doc.core_properties.title:
            doc.core_properties.title = metadata["title"]

    # Apply header & footer to the correct section
    target_section = 1 if has_cover else 0
    header_config = content_json.get("header")
    footer_config = content_json.get("footer")

    if header_config and target_section < len(doc.sections):
        _apply_header(doc, target_section, theme, header_config)
    if footer_config and target_section < len(doc.sections):
        _apply_footer(doc, target_section, theme, footer_config)

    # Force fields (like TOC) to update automatically on open
    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        update_fields = OxmlElement('w:updateFields')
        update_fields.set(qn('w:val'), 'true')
        doc.settings.element.append(update_fields)
    except Exception as e:
        print(f"Warning: Failed to inject updateFields setting: {e}", file=sys.stderr)

    return doc


# ============================================================================
# SECTION 8: QUALITY VALIDATION
# ============================================================================

def _validate_content(content_data):
    """Validate content JSON and return (errors, warnings) lists."""
    content_json = _normalize_content(content_data)
    errors = []
    warnings = []

    blocks = content_json.get("content", [])
    if not blocks:
        warnings.append("Content array is empty — document will have no content.")
        return errors, warnings

    # Check block types
    headings_seen = set()
    for i, block in enumerate(blocks):
        btype = block.get("type")
        if not btype:
            errors.append(f"Block {i}: missing 'type' field.")
            continue
        if btype not in SUPPORTED_BLOCK_TYPES:
            errors.append(f"Block {i}: unknown type '{btype}'. "
                          f"Supported: {', '.join(sorted(SUPPORTED_BLOCK_TYPES))}")
            continue

        # Type-specific validation
        if btype == "heading":
            level = block.get("level", 1)
            try:
                level = int(level)
            except (TypeError, ValueError):
                errors.append(f"Block {i}: heading level must be an integer.")
                continue
            text = block.get("text", "")
            if not text:
                warnings.append(f"Block {i}: heading has empty text.")
            if level < 1 or level > 4:
                errors.append(f"Block {i}: heading level {level} out of range (1-4).")
            key = f"H{level}:{text}"
            if key in headings_seen:
                warnings.append(f"Block {i}: duplicate heading '{text}' (level {level}).")
            headings_seen.add(key)

        elif btype == "paragraph":
            if not block.get("text"):
                warnings.append(f"Block {i}: paragraph has empty text.")

        elif btype in ("bullet_list", "numbered_list", "checklist"):
            items = block.get("items", [])
            if not items:
                warnings.append(f"Block {i}: {btype} has no items.")

        elif btype == "table":
            headers = block.get("headers", [])
            rows = block.get("rows", [])
            if not headers:
                errors.append(f"Block {i}: table has no headers.")
            for j, row in enumerate(rows):
                if len(row) != len(headers):
                    warnings.append(f"Block {i}: table row {j} has {len(row)} cells "
                                    f"but header has {len(headers)} columns.")

        elif btype == "cover_page":
            if not block.get("title"):
                errors.append(f"Block {i}: cover_page requires a 'title' field.")
            if i != 0:
                warnings.append(f"Block {i}: cover_page should be the first block.")

        elif btype == "timeline":
            items = block.get("items", [])
            for j, item in enumerate(items):
                if not item.get("phase"):
                    warnings.append(f"Block {i}: timeline item {j} missing 'phase'.")
                if not item.get("date"):
                    warnings.append(f"Block {i}: timeline item {j} missing 'date'.")

        elif btype in ("note", "warning", "important", "executive_summary"):
            if not block.get("text"):
                warnings.append(f"Block {i}: {btype} has empty text.")

        elif btype == "code":
            if not block.get("text"):
                warnings.append(f"Block {i}: code block has empty text.")

    # Theme validation
    theme_name = content_json.get("theme")
    if theme_name and theme_name not in THEMES:
        errors.append(f"Unknown theme '{theme_name}'. Available: {', '.join(THEMES.keys())}")

    return errors, warnings


# ============================================================================
# SECTION 9: ADVANCED EDITING
# ============================================================================

def _find_heading_in_body(doc, heading_text):
    """Find a heading element in the body by text. Returns (index, style_val) or (None, None)."""
    body = doc.element.body
    children = list(body)
    for i, element in enumerate(children):
        tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
        if tag != "p":
            continue
        pPr = element.find(qn("w:pPr"))
        if pPr is None:
            continue
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None:
            continue
        style_val = pStyle.get(qn("w:val"))
        if not style_val or not style_val.startswith("Heading"):
            continue
        # Gather text from runs
        full_text = ""
        for r in element.findall(qn("w:r")):
            t = r.find(qn("w:t"))
            if t is not None and t.text:
                full_text += t.text
        if full_text.strip().lower() == heading_text.strip().lower():
            return i, style_val
    return None, None


def _get_section_end(doc, heading_idx, heading_style):
    """Find the end index of a section (next same-or-higher level heading, or end of body)."""
    body = doc.element.body
    children = list(body)
    try:
        heading_level = int(heading_style.replace("Heading", "").strip())
    except ValueError:
        heading_level = 1

    for i in range(heading_idx + 1, len(children)):
        element = children[i]
        tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
        if tag != "p":
            continue
        pPr = element.find(qn("w:pPr"))
        if pPr is None:
            continue
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None:
            continue
        style_val = pStyle.get(qn("w:val"))
        if not style_val or not style_val.startswith("Heading"):
            continue
        try:
            level = int(style_val.replace("Heading", "").strip())
        except ValueError:
            continue
        if level <= heading_level:
            return i
    return len(children)


def _insert_elements_after(doc, body_index, theme, content_blocks):
    """Insert rendered content blocks after the element at body_index."""
    temp_doc = Document()
    _apply_default_styles(temp_doc, theme)
    _process_content_blocks(temp_doc, theme, content_blocks)

    body = doc.element.body
    temp_body = temp_doc.element.body

    new_elements = []
    for elem in list(temp_body):
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag != "sectPr":
            new_elements.append(elem)

    ref = list(body)[body_index]
    for elem in reversed(new_elements):
        ref.addnext(elem)


# ============================================================================
# SECTION 10: FILE I/O HELPERS
# ============================================================================

def _load_content_json(content_path):
    """Load and validate a content JSON file."""
    if not os.path.isfile(content_path):
        print(f"Error: Content file not found: {content_path}", file=sys.stderr)
        sys.exit(1)
    try:
        with open(content_path, "r", encoding="utf-8") as f:
            content = json.load(f)
    except json.JSONDecodeError as e:
        print(f"Error: Invalid content JSON — {e}", file=sys.stderr)
        sys.exit(1)
    return content


def _load_document(file_path):
    """Load an existing .docx file with error handling."""
    if not os.path.isfile(file_path):
        print(f"Error: File not found: {file_path}", file=sys.stderr)
        sys.exit(1)
    try:
        return Document(file_path)
    except Exception as e:
        print(f"Error: Cannot read file — it may be corrupted or not a valid .docx: {e}",
              file=sys.stderr)
        sys.exit(1)


# ============================================================================
# SECTION 11: SUBCOMMAND HANDLERS
# ============================================================================

def _convert_to_pdf(docx_path, pdf_path=None):
    """Convert .docx to .pdf using best available backend."""
    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    
    # Try docx2pdf (requires MS Word)
    try:
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        return pdf_path
    except ImportError:
        pass
    except Exception as e:
        print(f"Warning: docx2pdf failed ({e}), trying LibreOffice...", file=sys.stderr)
    
    # Try LibreOffice headless
    soffice = shutil.which("soffice")
    if soffice:
        out_dir = os.path.dirname(pdf_path) or "."
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", out_dir, docx_path],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0:
            generated = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
            generated_path = os.path.join(out_dir, generated)
            if generated_path != pdf_path:
                os.rename(generated_path, pdf_path)
            return pdf_path
    
    print("Error: No PDF backend available.\n"
          "  Install docx2pdf: pip install docx2pdf (requires MS Word)\n"
          "  Or install LibreOffice: https://www.libreoffice.org/", file=sys.stderr)
    sys.exit(1)


def cmd_create(args):
    """Handle the 'create' subcommand."""
    content_data = _load_content_json(args.content)
    
    if hasattr(args, "variables") and args.variables:
        if not os.path.isfile(args.variables):
            print(f"Error: Variables file not found: {args.variables}", file=sys.stderr)
            sys.exit(1)
        try:
            with open(args.variables, "r", encoding="utf-8") as f:
                if args.variables.endswith(".csv"):
                    reader = csv.DictReader(f)
                    variables = next(reader)
                else:
                    variables = json.load(f)
            content_data = _substitute_variables(content_data, variables)
        except Exception as e:
            print(f"Error loading variables: {e}", file=sys.stderr)
            sys.exit(1)

    doc = _build_document(content_data, theme_override=args.theme)

    output_dir = args.output_dir or DEFAULT_OUTPUT_DIR
    os.makedirs(output_dir, exist_ok=True)

    title = args.title
    for ext in (".docx", ".doc", ".pdf"):
        if title.lower().endswith(ext):
            title = title[:-len(ext)]
            break

    output_path = os.path.join(output_dir, f"{title}.docx")
    doc.save(output_path)
    print(f"Success! Document created: {output_path}")

    if args.output_format in ("pdf", "both"):
        pdf_path = os.path.join(output_dir, f"{title}.pdf")
        print(f"Converting to PDF...")
        _convert_to_pdf(output_path, pdf_path)
        print(f"Success! PDF created: {pdf_path}")
        
        if args.output_format == "pdf" and not args.keep_docx:
            os.remove(output_path)


def cmd_edit(args):
    """Handle the 'edit' subcommand with multiple action types."""
    doc = _load_document(args.file)
    theme = _resolve_theme(args.theme)

    if args.action == "append":
        if not args.content:
            print("Error: --content is required for 'append' action.", file=sys.stderr)
            sys.exit(1)
        content_data = _load_content_json(args.content)
        blocks = content_data if isinstance(content_data, list) else content_data.get("content", content_data)
        if isinstance(blocks, dict) and "content" not in blocks:
            blocks = [blocks]
        if not isinstance(blocks, list):
            blocks = []
        _process_content_blocks(doc, theme, blocks)
        doc.save(args.file)
        print(f"Success! Content appended to: {args.file}")

    elif args.action == "replace":
        if not args.find or args.replace_with is None:
            print("Error: --find and --replace-with are required for 'replace' action.",
                  file=sys.stderr)
            sys.exit(1)
        find_text = args.find
        replace_text = args.replace_with
        count = 0
        for paragraph in doc.paragraphs:
            count += _replace_in_paragraph(paragraph, find_text, replace_text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        count += _replace_in_paragraph(paragraph, find_text, replace_text)
        if count == 0:
            print(f"Warning: Text '{find_text}' not found in the document.", file=sys.stderr)
        else:
            doc.save(args.file)
            print(f"Success! Replaced {count} occurrence(s) in: {args.file}")

    elif args.action == "insert-after":
        if not args.heading or not args.content:
            print("Error: --heading and --content required for 'insert-after'.", file=sys.stderr)
            sys.exit(1)
        content_data = _load_content_json(args.content)
        blocks = content_data if isinstance(content_data, list) else content_data.get("content", content_data)
        if not isinstance(blocks, list):
            blocks = [blocks]
        idx, _ = _find_heading_in_body(doc, args.heading)
        if idx is None:
            print(f"Error: Heading '{args.heading}' not found in document.", file=sys.stderr)
            sys.exit(1)
        _insert_elements_after(doc, idx, theme, blocks)
        doc.save(args.file)
        print(f"Success! Content inserted after '{args.heading}' in: {args.file}")

    elif args.action == "replace-section":
        if not args.heading or not args.content:
            print("Error: --heading and --content required for 'replace-section'.", file=sys.stderr)
            sys.exit(1)
        content_data = _load_content_json(args.content)
        blocks = content_data if isinstance(content_data, list) else content_data.get("content", content_data)
        if not isinstance(blocks, list):
            blocks = [blocks]
        idx, style_val = _find_heading_in_body(doc, args.heading)
        if idx is None:
            print(f"Error: Heading '{args.heading}' not found.", file=sys.stderr)
            sys.exit(1)
        section_end = _get_section_end(doc, idx, style_val)
        # Remove old section content (keep the heading)
        body = doc.element.body
        children = list(body)
        for i in range(section_end - 1, idx, -1):
            body.remove(children[i])
        # Insert new content after heading
        _insert_elements_after(doc, idx, theme, blocks)
        doc.save(args.file)
        print(f"Success! Section '{args.heading}' replaced in: {args.file}")

    elif args.action == "delete-section":
        if not args.heading:
            print("Error: --heading required for 'delete-section'.", file=sys.stderr)
            sys.exit(1)
        idx, style_val = _find_heading_in_body(doc, args.heading)
        if idx is None:
            print(f"Error: Heading '{args.heading}' not found.", file=sys.stderr)
            sys.exit(1)
        section_end = _get_section_end(doc, idx, style_val)
        body = doc.element.body
        children = list(body)
        for i in range(section_end - 1, idx - 1, -1):
            body.remove(children[i])
        doc.save(args.file)
        print(f"Success! Section '{args.heading}' deleted from: {args.file}")

    elif args.action == "update-table":
        if args.table_index is None or not args.table_data:
            print("Error: --table-index and --table-data required for 'update-table'.",
                  file=sys.stderr)
            sys.exit(1)
        tables = [t for t in doc.tables if not (len(t.rows) == 1 and len(t.columns) == 1)]
        if args.table_index < 0 or args.table_index >= len(tables):
            print(f"Error: Table index {args.table_index} out of range. "
                  f"Document has {len(tables)} table(s).", file=sys.stderr)
            sys.exit(1)
        # Load new table data
        if not os.path.isfile(args.table_data):
            print(f"Error: Table data file not found: {args.table_data}", file=sys.stderr)
            sys.exit(1)
        try:
            with open(args.table_data, "r", encoding="utf-8") as f:
                td = json.load(f)
        except json.JSONDecodeError as e:
            print(f"Error: Invalid table data JSON — {e}", file=sys.stderr)
            sys.exit(1)
        headers = td.get("headers", [])
        rows = td.get("rows", [])

        # Create new table in temp doc
        temp_doc = Document()
        _apply_default_styles(temp_doc, theme)
        _add_table_block(temp_doc, theme, headers, rows)

        # Find new table element
        new_tbl = None
        for elem in list(temp_doc.element.body):
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag == "tbl":
                new_tbl = elem
                break
        if new_tbl is None:
            print("Error: Failed to create replacement table.", file=sys.stderr)
            sys.exit(1)

        # Swap: replace old table
        target_tbl = tables[args.table_index]._tbl
        p = target_tbl.getparent()
        p.replace(target_tbl, new_tbl)

        doc.save(args.file)
        print(f"Success! Table {args.table_index} updated in: {args.file}")

    else:
        print(f"Error: Unknown action '{args.action}'.", file=sys.stderr)
        sys.exit(1)


def cmd_add_table(args):
    """Handle the 'add-table' subcommand."""
    doc = _load_document(args.file)
    theme = _resolve_theme(args.theme)

    if args.table_data:
        if not os.path.isfile(args.table_data):
            print(f"Error: Table data file not found: {args.table_data}", file=sys.stderr)
            sys.exit(1)
        try:
            with open(args.table_data, "r", encoding="utf-8") as f:
                td = json.load(f)
        except json.JSONDecodeError as e:
            print(f"Error: Invalid table data JSON — {e}", file=sys.stderr)
            sys.exit(1)
        headers = td.get("headers", [])
        rows = td.get("rows", [])
    elif args.headers and args.rows:
        try:
            headers = json.loads(args.headers)
        except json.JSONDecodeError as e:
            print(f"Error: Invalid --headers JSON — {e}", file=sys.stderr)
            sys.exit(1)
        try:
            rows = json.loads(args.rows)
        except json.JSONDecodeError as e:
            print(f"Error: Invalid --rows JSON — {e}", file=sys.stderr)
            sys.exit(1)
    else:
        print("Error: Provide either --table-data or both --headers and --rows.",
              file=sys.stderr)
        sys.exit(1)

    if not isinstance(headers, list):
        print("Error: headers must be a JSON array of strings.", file=sys.stderr)
        sys.exit(1)
    if not isinstance(rows, list):
        print("Error: rows must be a JSON array of arrays.", file=sys.stderr)
        sys.exit(1)

    _add_table_block(doc, theme, headers, rows)
    doc.save(args.file)
    print(f"Success! Table added to: {args.file}")


def cmd_remove_table(args):
    """Handle the 'remove-table' subcommand."""
    doc = _load_document(args.file)
    tables = [t for t in doc.tables if not (len(t.rows) == 1 and len(t.columns) == 1)]
    if args.index < 0 or args.index >= len(tables):
        print(f"Error: Table index {args.index} out of range.", file=sys.stderr)
        sys.exit(1)
    table_element = tables[args.index]._tbl
    table_element.getparent().remove(table_element)
    doc.save(args.file)
    print(f"Success! Table at index {args.index} removed from: {args.file}")


def cmd_info(args):
    """Handle the 'info' subcommand — output document structure as JSON."""
    doc = _load_document(args.file)

    content_list = []
    idx = 0

    for element in doc.element.body:
        tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag

        if tag == "p":
            pPr = element.find(qn("w:pPr"))
            pStyle = None
            if pPr is not None:
                pStyleEl = pPr.find(qn("w:pStyle"))
                if pStyleEl is not None:
                    pStyle = pStyleEl.get(qn("w:val"))

            # Gather run text
            full_text = ""
            for r in element.findall(qn("w:r")):
                t = r.find(qn("w:t"))
                if t is not None and t.text:
                    full_text += t.text

            if pStyle and pStyle.startswith("Heading"):
                try:
                    level = int(pStyle.replace("Heading", "").strip())
                except ValueError:
                    level = 1
                content_list.append({
                    "index": idx, "type": "heading", "level": level, "text": full_text,
                })
            elif pStyle in ("ListBullet", "List Bullet"):
                content_list.append({"index": idx, "type": "bullet_item", "text": full_text})
            elif pStyle in ("ListNumber", "List Number"):
                content_list.append({"index": idx, "type": "numbered_item", "text": full_text})
            else:
                if full_text.strip():
                    content_list.append({"index": idx, "type": "paragraph", "text": full_text})
            idx += 1

        elif tag == "tbl":
            tbl_rows = element.findall(qn("w:tr"))
            cols = 0
            if tbl_rows:
                cols = len(tbl_rows[0].findall(qn("w:tc")))
                
            if len(tbl_rows) == 1 and cols == 1:
                idx += 1
                continue
                
            content_list.append({
                "index": idx, "type": "table", "rows": len(tbl_rows), "cols": cols,
            })
            idx += 1

    # Metadata
    props = doc.core_properties
    metadata = {}
    if props.author:
        metadata["author"] = props.author
    if props.title:
        metadata["title"] = props.title
    if props.version:
        metadata["version"] = props.version
    if props.content_status:
        metadata["status"] = props.content_status
    if props.category:
        metadata["department"] = props.category

    table_count = sum(1 for c in content_list if c["type"] == "table")
    para_count = sum(1 for c in content_list if c["type"] != "table")

    result = {
        "filename": os.path.basename(args.file),
        "paragraphs": para_count,
        "tables": table_count,
        "sections": len(doc.sections),
        "metadata": metadata if metadata else None,
        "content": content_list,
    }

    try:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2)
        print(f"Success! Document info written to: {args.output}")
    except OSError as e:
        print(f"Error writing to file {args.output}: {e}", file=sys.stderr)
        sys.exit(1)


def cmd_validate(args):
    """Handle the 'validate' subcommand."""
    content_data = _load_content_json(args.content)
    errors, warnings = _validate_content(content_data)

    for w in warnings:
        print(f"  WARNING: {w}", file=sys.stderr)
    for e in errors:
        print(f"  ERROR: {e}", file=sys.stderr)

    if errors:
        print(f"Invalid — {len(errors)} error(s), {len(warnings)} warning(s).")
        sys.exit(1)
    elif warnings:
        print(f"Valid with {len(warnings)} warning(s).")
    else:
        print("Valid — no issues found.")


def cmd_batch(args):
    """Handle the 'batch' subcommand."""
    if not os.path.isfile(args.variables):
        print(f"Error: Variables CSV file not found: {args.variables}", file=sys.stderr)
        sys.exit(1)
        
    template_data = _load_content_json(args.template)
    output_dir = args.output_dir or DEFAULT_OUTPUT_DIR
    os.makedirs(output_dir, exist_ok=True)
    
    success_count = 0
    error_count = 0
    
    try:
        with open(args.variables, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            
            # Require _output_name column
            if "_output_name" not in reader.fieldnames:
                print("Error: CSV must contain an '_output_name' column", file=sys.stderr)
                sys.exit(1)
                
            for i, row in enumerate(reader):
                try:
                    output_name = row.pop("_output_name")
                    content_data = _substitute_variables(template_data, row)
                    doc = _build_document(content_data, theme_override=args.theme)
                    
                    title = output_name
                    for ext in (".docx", ".doc", ".pdf"):
                        if title.lower().endswith(ext):
                            title = title[:-len(ext)]
                            break
                            
                    output_path = os.path.join(output_dir, f"{title}.docx")
                    doc.save(output_path)
                    print(f"Created: {output_path}")
                    
                    if args.output_format in ("pdf", "both"):
                        pdf_path = os.path.join(output_dir, f"{title}.pdf")
                        _convert_to_pdf(output_path, pdf_path)
                        print(f"Created: {pdf_path}")
                        if args.output_format == "pdf" and not args.keep_docx:
                            os.remove(output_path)
                            
                    success_count += 1
                except Exception as e:
                    print(f"Error processing row {i+1} ({output_name}): {e}", file=sys.stderr)
                    error_count += 1
                    
    except Exception as e:
        print(f"Error reading CSV: {e}", file=sys.stderr)
        sys.exit(1)
        
    print(f"\nBatch complete: {success_count} success, {error_count} errors.")


# ============================================================================
# SECTION 12: CLI ENTRYPOINT
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Doc Manager v2.0 — Professional AI Document Generation Tool",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # --- create ---
    p_create = subparsers.add_parser("create", help="Create a new .docx document")
    p_create.add_argument("--title", required=True,
                          help="Document filename (without .docx extension)")
    p_create.add_argument("--content", required=True,
                          help="Path to JSON file describing document content")
    p_create.add_argument("--theme", default=None,
                          choices=list(THEMES.keys()),
                          help="Document theme (default: professional)")
    p_create.add_argument("--output-dir", default=None,
                          help=f"Output directory (default: {DEFAULT_OUTPUT_DIR})")
    p_create.add_argument("--output-format", default="docx",
                          choices=["docx", "pdf", "both"],
                          help="Output format (default: docx)")
    p_create.add_argument("--keep-docx", action="store_true",
                          help="Keep intermediate .docx file when output format is pdf")
    p_create.add_argument("--variables", default=None,
                          help="Path to JSON or CSV variables file for template substitution")

    # --- batch ---
    p_batch = subparsers.add_parser("batch", help="Generate multiple documents from a template and CSV")
    p_batch.add_argument("--template", required=True,
                         help="Path to JSON template file")
    p_batch.add_argument("--variables", required=True,
                         help="Path to CSV variables file (must include _output_name column)")
    p_batch.add_argument("--theme", default=None,
                         choices=list(THEMES.keys()),
                         help="Document theme (default: professional)")
    p_batch.add_argument("--output-dir", default=None,
                         help=f"Output directory (default: {DEFAULT_OUTPUT_DIR})")
    p_batch.add_argument("--output-format", default="docx",
                         choices=["docx", "pdf", "both"],
                         help="Output format (default: docx)")
    p_batch.add_argument("--keep-docx", action="store_true",
                         help="Keep intermediate .docx file when output format is pdf")

    # --- edit ---
    p_edit = subparsers.add_parser("edit", help="Edit an existing .docx document")
    p_edit.add_argument("--file", required=True, help="Path to existing .docx file")
    p_edit.add_argument("--action", required=True,
                        choices=["append", "replace", "insert-after",
                                 "replace-section", "delete-section", "update-table"],
                        help="Edit action type")
    p_edit.add_argument("--content", default=None,
                        help="Path to JSON file with content (for append/insert-after/replace-section)")
    p_edit.add_argument("--find", default=None, help="Text to find (for replace)")
    p_edit.add_argument("--replace-with", default=None, help="Replacement text (for replace)")
    p_edit.add_argument("--heading", default=None,
                        help="Target heading (for insert-after/replace-section/delete-section)")
    p_edit.add_argument("--table-index", type=int, default=None,
                        help="Table index (for update-table)")
    p_edit.add_argument("--table-data", default=None,
                        help="Path to table data JSON (for update-table)")
    p_edit.add_argument("--theme", default="professional",
                        choices=list(THEMES.keys()),
                        help="Theme for new content (default: professional)")

    # --- add-table ---
    p_add_table = subparsers.add_parser("add-table", help="Add a table to a document")
    p_add_table.add_argument("--file", required=True, help="Path to existing .docx file")
    p_add_table.add_argument("--table-data", default=None,
                             help='Path to JSON file with {"headers": [...], "rows": [[...]]}')
    p_add_table.add_argument("--headers", default=None,
                             help="JSON array of column headers (alt to --table-data)")
    p_add_table.add_argument("--rows", default=None,
                             help="JSON array of row arrays (alt to --table-data)")
    p_add_table.add_argument("--theme", default="professional",
                             choices=list(THEMES.keys()),
                             help="Table theme (default: professional)")

    # --- remove-table ---
    p_remove = subparsers.add_parser("remove-table",
                                     help="Remove a table from a document by index")
    p_remove.add_argument("--file", required=True, help="Path to existing .docx file")
    p_remove.add_argument("--index", required=True, type=int,
                          help="0-based index of the table to remove")

    # --- info ---
    p_info = subparsers.add_parser("info",
                                   help="Show structure and content of a .docx document")
    p_info.add_argument("--file", required=True, help="Path to existing .docx file")
    p_info.add_argument("--output", required=True, help="Output JSON file path")

    # --- validate ---
    p_validate = subparsers.add_parser("validate",
                                       help="Validate a content JSON file before creation")
    p_validate.add_argument("--content", required=True,
                            help="Path to content JSON file to validate")

    args = parser.parse_args()

    if args.command == "create":
        cmd_create(args)
    elif args.command == "batch":
        cmd_batch(args)
    elif args.command == "edit":
        cmd_edit(args)
    elif args.command == "add-table":
        cmd_add_table(args)
    elif args.command == "remove-table":
        cmd_remove_table(args)
    elif args.command == "info":
        cmd_info(args)
    elif args.command == "validate":
        cmd_validate(args)
    else:
        print(f"Unknown command: {args.command}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
