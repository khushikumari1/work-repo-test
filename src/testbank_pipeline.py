"""
testbank_pipeline.py — Cengage test bank audit pipeline (single-file build).

Organised into four sections:
    Section 1 — Shared Data Models   : TestBankItem, OutputItem, ChunkStats, constants
    Section 2 — DOCX I/O             : Doc conversion, chapter extraction, test bank
                                        parsing, docx output builder
    Section 3 — Gemini Engine         : Gemini API client, prompt builder, response
                                        parser, chunked processor, content validator
    Section 4 — Pipeline Orchestrator : TestBankPipeline (wires everything together),
                                        cost/output logging, CLI entry point

Ground-truth document structure (TESTBANK_CH01.docx):
    Each question is a 1x1 "shell" table in the document body.  Inside the
    single cell: an optional [REVISE] paragraph, the stem paragraph, a nested
    3-column options table (MCQ only), a nested 2-column metadata table, and
    the [AI REVIEW] block.  Tracked changes are faked with run-level formatting
    (strikethrough dark-red for deletions, bold green for insertions) rather
    than real <w:ins>/<w:del> markup.
"""

# =====================================================================
# Standard library imports
# =====================================================================
import html
import logging
import os
import platform
import re
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape

# =====================================================================
# Third-party imports
# =====================================================================
from google import genai
from google.genai import types
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table


# =====================================================================
# =====================================================================
#  SECTION 1 — SHARED DATA MODELS
# =====================================================================
# =====================================================================

# The canonical metadata field order used everywhere: parsing, XML
# serialization, XML de-serialization, validation, and docx writing.
# Every section that needs "the list of metadata labels" references this
# instead of re-typing it.
METADATA_FIELDS: List[str] = [
    "answer",
    "points",
    "difficulty",
    "references",
    "learning_objectives",
    "accrediting_standards",
    "keywords",
]

# Maps a metadata field name to the literal label text as it appears in the
# source .docx (including the non-breaking spaces Word uses after the colon).
NBSP = "\xa0"
METADATA_LABELS: Dict[str, str] = {
    "answer": f"ANSWER:{NBSP}{NBSP}",
    "points": f"POINTS:{NBSP}{NBSP}",
    "difficulty": f"DIFFICULTY:{NBSP}{NBSP}",
    "references": f"REFERENCES:{NBSP}{NBSP}",
    "learning_objectives": f"LEARNING{NBSP}OBJECTIVES:{NBSP}{NBSP}",
    "accrediting_standards": f"ACCREDITING{NBSP}STANDARDS:{NBSP}{NBSP}",
    "keywords": f"KEYWORDS:{NBSP}{NBSP}",
    "other": f"OTHER:{NBSP}{NBSP}",
}

OPTION_LETTERS: List[str] = ["a", "b", "c", "d"]

DECISION_KEEP = "KEEP"
DECISION_REVISE = "REVISE"
DECISION_REMOVE = "REMOVE"
DECISION_PARSE_ERROR = "PARSE_ERROR"
VALID_DECISIONS = {DECISION_KEEP, DECISION_REVISE, DECISION_REMOVE, DECISION_PARSE_ERROR}


@dataclass
class TestBankItem:
    """A single question extracted from the previous-edition test bank."""

    number: int
    has_revise_marker: bool = False
    stem: str = ""
    options: Dict[str, str] = field(default_factory=dict)  # {'a': ..., 'b': ..., ...}
    answer: str = ""
    points: str = ""
    difficulty: str = ""
    references: str = ""
    learning_objectives: str = ""
    accrediting_standards: str = ""
    keywords: str = ""
    other_html: str = ""
    is_essay: bool = False  # True when there are no a/b/c/d options

    def metadata(self) -> Dict[str, str]:
        """Return the metadata fields as a plain dict, in canonical order."""
        return {f: getattr(self, f) for f in METADATA_FIELDS}


@dataclass
class OutputItem:
    """A single question after AI review, ready to be written to the output docx."""

    number: int
    decision: str = DECISION_KEEP
    stem: str = ""
    options: Dict[str, str] = field(default_factory=dict)
    answer: str = ""
    points: str = ""
    difficulty: str = ""
    references: str = ""
    learning_objectives: str = ""
    accrediting_standards: str = ""
    keywords: str = ""
    other_html: str = ""
    ai_review_text: str = ""
    revisions: str = ""
    is_essay: bool = False
    flags: List[str] = field(default_factory=list)  # ContentValidator annotations

    def metadata(self) -> Dict[str, str]:
        return {f: getattr(self, f) for f in METADATA_FIELDS}

    @classmethod
    def from_test_bank_item(cls, item: TestBankItem, decision: str = DECISION_PARSE_ERROR,
                             ai_review_text: str = "") -> "OutputItem":
        """Fallback constructor: build an OutputItem straight from the original
        TestBankItem when Gemini's response for this item couldn't be parsed.
        Guarantees we never silently drop an item."""
        return cls(
            number=item.number,
            decision=decision,
            stem=item.stem,
            options=dict(item.options),
            answer=item.answer,
            points=item.points,
            difficulty=item.difficulty,
            references=item.references,
            learning_objectives=item.learning_objectives,
            accrediting_standards=item.accrediting_standards,
            keywords=item.keywords,
            other_html=item.other_html,
            ai_review_text=ai_review_text,
            is_essay=item.is_essay,
        )


@dataclass
class ChunkStats:
    """Per-chunk processing stats, used for logging and the Streamlit progress UI."""

    chunk_index: int
    items_sent: int
    items_returned: int
    elapsed_seconds: float
    retried: bool = False
    # Token counts returned by the Gemini API for this chunk
    input_tokens: int = 0
    output_tokens: int = 0

    @property
    def total_tokens(self) -> int:
        return self.input_tokens + self.output_tokens


# =====================================================================
# =====================================================================
#  SECTION 2 — DOCX I/O
# =====================================================================
# =====================================================================

logger = logging.getLogger(__name__)


# =====================================================================
# Doc conversion (.doc -> .docx) and PDF text extraction
# =====================================================================

class DocConversionError(RuntimeError):
    """Raised when a .doc file cannot be converted to .docx by any available method."""


def convert_doc_to_docx(doc_path: str) -> str:
    """Convert a legacy .doc file to .docx using LibreOffice (headless), with a
    Word-COM fallback on Windows. Returns the path to the converted .docx file."""
    filename = os.path.splitext(os.path.basename(doc_path))[0]
    output_dir = tempfile.gettempdir()
    docx_path = os.path.join(output_dir, filename + ".docx")

    if os.path.exists(docx_path):
        logger.info("Using existing converted file: %s", docx_path)
        return docx_path

    system = platform.system()
    if system == "Windows":
        soffice_candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "soffice.exe", "soffice",
        ]
    elif system == "Darwin":
        soffice_candidates = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice", "soffice", "libreoffice",
        ]
    else:
        soffice_candidates = ["/usr/bin/soffice", "/usr/bin/libreoffice", "soffice", "libreoffice"]

    for soffice in soffice_candidates:
        try:
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "docx", "--outdir", output_dir, doc_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=60,
            )
            if result.returncode == 0 and os.path.exists(docx_path):
                logger.info("Converted %s -> %s via %s", doc_path, docx_path, soffice)
                return docx_path
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
        except Exception as exc:  # noqa: BLE001 - best-effort fallback chain
            logger.debug("Conversion attempt failed with %s: %s", soffice, exc)
            continue

    if system == "Windows":
        converted = _convert_via_word_com(doc_path, docx_path)
        if converted:
            return converted

    raise DocConversionError(
        "Could not convert .doc file to .docx format. Please install LibreOffice:\n"
        "  - Windows: https://www.libreoffice.org/download/download/\n"
        "  - Linux: sudo apt-get install libreoffice\n"
        "  - macOS: brew install --cask libreoffice\n"
        "Or convert the file manually to .docx format."
    )


def _convert_via_word_com(doc_path: str, docx_path: str) -> str:
    """Windows-only fallback using the Word COM automation API."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        logger.debug("pythoncom/win32com not available; skipping Word COM fallback")
        return ""

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(FileName=os.path.abspath(doc_path), ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs2(os.path.abspath(docx_path), FileFormat=16)  # wdFormatDocumentDefault
        logger.info("Converted via Word COM: %s", docx_path)
        return docx_path
    except Exception as exc:  # noqa: BLE001
        logger.debug("Word COM conversion failed: %s", exc)
        return ""
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def ensure_docx(file_path: str) -> str:
    """Given a .doc or .docx path, return a path guaranteed to be .docx,
    converting if necessary."""
    lower = file_path.lower()
    if lower.endswith(".docx"):
        return file_path
    if lower.endswith(".doc"):
        return convert_doc_to_docx(file_path)
    raise ValueError(f"Not a Word document: {file_path}")


def read_pdf_text(file_path: str) -> str:
    """Extract page-by-page text from a PDF using PyPDF2."""
    from PyPDF2 import PdfReader

    reader = PdfReader(file_path)
    pages = []
    for page in reader.pages:
        text = page.extract_text()
        if text and text.strip():
            pages.append(text)
    logger.info("Extracted text from %d pages in %s", len(reader.pages), file_path)
    return "\n".join(pages)


# =====================================================================
# Textbook Chapter Extractor
# =====================================================================

_HEADING_STYLE_PREFIX = "Heading"
_HEADING_MARKERS = {1: "#", 2: "##", 3: "###"}


class ChapterExtractor:
    """Extracts structured text from a textbook chapter file (.doc/.docx/.pdf)."""

    def extract(self, file_path: str) -> str:
        suffix = Path(file_path).suffix.lower()
        if suffix == ".pdf":
            text = read_pdf_text(file_path)
        elif suffix in (".doc", ".docx"):
            text = self._extract_docx(file_path)
        else:
            raise ValueError(f"Unsupported chapter file type: {suffix}")

        logger.info("Extracted %d characters from chapter file %s", len(text), file_path)
        return text

    def _extract_docx(self, file_path: str) -> str:
        docx_path = ensure_docx(file_path)
        document = Document(docx_path)
        lines = []
        for paragraph in document.paragraphs:
            text = self._resolved_text(paragraph)
            if not text:
                continue
            level = self._heading_level(paragraph.style.name if paragraph.style else "")
            if level:
                marker = _HEADING_MARKERS.get(level, "#" * level)
                lines.append(f"{marker} {text}")
            else:
                lines.append(text)
        return "\n\n".join(lines)

    @staticmethod
    def _resolved_text(paragraph) -> str:
        """Like paragraph.text, but resolves real Word tracked changes:
        keeps <w:ins> text, drops anything inside <w:del> (deleted/rejected
        text), regardless of whether it's stored as w:delText or w:t."""
        parts = []
        for node in paragraph._p.iter():
            tag = node.tag
            if tag.endswith("}delText"):
                continue
            if tag.endswith("}t"):
                if any(a.tag.endswith("}del") for a in node.iterancestors()):
                    continue
                parts.append(node.text or "")
        return "".join(parts).strip()

    @staticmethod
    def _heading_level(style_name: str):
        if not style_name.startswith(_HEADING_STYLE_PREFIX):
            return None
        try:
            return int(style_name.replace(_HEADING_STYLE_PREFIX, "").strip())
        except ValueError:
            return None

    @staticmethod
    def section_headers(structured_text: str, limit: int = 3):
        headers = [line for line in structured_text.split("\n\n") if line.startswith("#")]
        return headers[:limit]


# =====================================================================
# Structured Test Bank Parser (Cognero table format)
# =====================================================================

_REVISE_MARKER = "[REVISE]"
_STEM_PATTERN = re.compile(r"^(\d+)\." + NBSP)

_LABEL_MAP = {
    "ANSWER": "answer",
    "POINTS": "points",
    "DIFFICULTY": "difficulty",
    "REFERENCES": "references",
    "LEARNING OBJECTIVES": "learning_objectives",
    "LEARNING\xa0OBJECTIVES": "learning_objectives",
    "ACCREDITING STANDARDS": "accrediting_standards",
    "ACCREDITING\xa0STANDARDS": "accrediting_standards",
    "KEYWORDS": "keywords",
    "OTHER": "other_html",
}
_LABEL_CLEAN = re.compile(r"[\s\xa0:]+$")


def _normalise_label(raw: str) -> str:
    return _LABEL_CLEAN.sub("", raw).upper()


class TestBankParseError(RuntimeError):
    """Raised when the test bank structure can't be parsed as expected."""


class TestBankParser:
    """Parses a Cognero-format test bank .doc/.docx into TestBankItem objects.

    Each top-level table in the document is one question (see module
    docstring for the full cell layout)."""

    def parse(self, file_path: str) -> List[TestBankItem]:
        docx_path = ensure_docx(file_path)
        document = Document(docx_path)
        items = self._parse_document(document)
        logger.info("Parsed %d test bank items from %s", len(items), file_path)
        return items

    def _parse_document(self, document: Document) -> List[TestBankItem]:
        items: List[TestBankItem] = []
        for table in document.tables:
            item = self._parse_question_table(table)
            if item is not None:
                items.append(item)
        return items

    def _parse_question_table(self, table: Table) -> Optional[TestBankItem]:
        try:
            cell = table.cell(0, 0)
        except IndexError:
            return None

        has_revise = False
        stem = ""
        number = 0

        for para in cell.paragraphs:
            text = para.text
            if text.strip() == _REVISE_MARKER:
                has_revise = True
                continue
            m = _STEM_PATTERN.match(text)
            if m:
                number = int(m.group(1))
                stem = text
                break

        if not stem:
            return None

        item = TestBankItem(number=number, has_revise_marker=has_revise, stem=stem)

        nested = cell.tables  # tables nested inside this cell

        if len(nested) >= 1:
            self._parse_options_table(nested[0], item)
        if len(nested) >= 2:
            self._parse_metadata_table(nested[1], item)
        elif len(nested) == 1:
            first_cell_text = nested[0].cell(0, 0).text.strip()
            if _normalise_label(first_cell_text) in _LABEL_MAP:
                self._parse_metadata_table(nested[0], item)

        item.is_essay = not bool(item.options)
        return item

    @staticmethod
    def _parse_options_table(table: Table, item: TestBankItem) -> None:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 3:
                letter_raw = cells[0].text.strip().rstrip(".")
                text_val = cells[1].text.strip() if len(cells) > 1 else ""
            else:
                letter_raw = cells[1].text.strip().rstrip(".")
                text_val = cells[2].text.strip()
            letter = letter_raw.lower()
            if letter in ("a", "b", "c", "d", "e") and text_val:
                item.options[letter] = text_val

    @staticmethod
    def _parse_metadata_table(table: Table, item: TestBankItem) -> None:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            label_raw = cells[0].text.strip()
            value = cells[1].text.strip()
            key = _LABEL_MAP.get(_normalise_label(label_raw))
            if key is None:
                continue
            if key == "other_html":
                item.other_html = value
            else:
                setattr(item, key, value)

    def summarize(self, items: List[TestBankItem]) -> str:
        total = len(items)
        with_revise = sum(1 for it in items if it.has_revise_marker)
        with_other = sum(1 for it in items if it.other_html.strip())
        essays = sum(1 for it in items if it.is_essay)
        mcq = total - essays
        return (
            f"Total items: {total}\n"
            f"Items with [REVISE] marker: {with_revise}\n"
            f"Items with OTHER field: {with_other}\n"
            f"Multiple-choice items: {mcq}\n"
            f"Essay items: {essays}"
        )


# =====================================================================
# DOCX Output Document Builder
# =====================================================================
# Colors confirmed from TESTBANK_CH01.docx XML (see module docstring).
_COLOR_REVISE_TAG = RGBColor(0xE3, 0x6C, 0x09)   # orange   [REVISE]
_COLOR_REVIEW      = RGBColor(0x1F, 0x4E, 0x99)   # dark blue  [AI REVIEW] block
_COLOR_DEL         = RGBColor(0xC0, 0x00, 0x00)   # dark red   removed text
_COLOR_INS         = RGBColor(0x00, 0x80, 0x00)   # green      inserted text
_COLOR_BLACK       = RGBColor(0x00, 0x00, 0x00)

_NESTED_TABLE_STYLE = "questionMetaData"  # style name used by the source template

_REVIEW_LABELS = [
    "Textbook Verification / Alignment:",
    "Decision / Status:",
    "Distractor Breakdown:",
    "Alignment Verdict:",
    "Review Summary:",
    "REVISIONS:",
]
_DISTRACTOR_LINE = re.compile(r"^([a-dA-D])\.\s*(.*)$")
_DEL_INS_PATTERN = re.compile(r"<(DEL|INS)>(.*?)</\1>", re.DOTALL | re.IGNORECASE)


def render_tracked_runs(paragraph, text: str) -> None:
    """Split `text` on <DEL>...</DEL> / <INS>...</INS> markers and add runs
    with the corresponding tracked-change-style formatting. Plain text
    outside any marker is added unformatted (default black)."""
    pos = 0
    for m in _DEL_INS_PATTERN.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        tag, inner = m.group(1).upper(), m.group(2)
        run = paragraph.add_run(inner)
        if tag == "DEL":
            run.font.strike = True
            run.font.color.rgb = _COLOR_DEL
        else:
            run.bold = True
            run.font.color.rgb = _COLOR_INS
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])

def render_fully_struck(paragraph, text: str) -> None:
    """Render an entire line as strikethrough dark-red text — used for the
    stem and options of REMOVE items, so the whole item reads as visually
    'crossed out' at a glance, matching the target template."""
    run = paragraph.add_run(strip_tracked_markup(text))
    run.font.strike = True
    run.font.color.rgb = _COLOR_DEL

def strip_tracked_markup(text: str) -> str:
    """Plain-text version of a tracked field — used when parsing input items
    (which never contain markup) and for anywhere we need the literal
    final text (e.g. length checks in ContentValidator)."""
    return _DEL_INS_PATTERN.sub(lambda m: m.group(2), text or "")


def _clear_default_paragraph(cell) -> None:
    for paragraph in list(cell.paragraphs):
        paragraph._element.getparent().remove(paragraph._element)

def _set_column_widths(table, widths) -> None:
    """Force fixed column widths instead of autofit, so narrow letter/
    checkbox columns don't get stretched even wide by the table style."""
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)

    old_grid = table._tbl.find(qn("w:tblGrid"))
    if old_grid is not None:
        table._tbl.remove(old_grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(w))
        grid.append(col)
    tbl_pr.addnext(grid)

    for row in table.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                row.cells[i].width = w

def _apply_borderless(table) -> None:
    """No visible grid lines — matches the source template's nested tables,
    which rely on the 'questionMetaData' style rather than explicit borders."""
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge = OxmlElement(f"w:{side}")
        edge.set(qn("w:val"), "none")
        borders.append(edge)
    tbl_pr.append(borders)


def _try_apply_style(table, style_name: str) -> None:
    """Best-effort: use the template's real table style if present, else
    fall back to a plain borderless table so output still looks reasonable
    even for inputs that didn't carry that style (e.g. some .doc conversions)."""
    try:
        table.style = style_name
    except (KeyError, ValueError):
        _apply_borderless(table)


def _match_review_label(line: str) -> Optional[Tuple[str, str]]:
    for label in _REVIEW_LABELS:
        if line.startswith(label):
            return label, line[len(label):]
    return None


class DocxBuilder:
    """Builds the audited test bank .docx from a list of OutputItems,
    reproducing the exact table-in-table shape and color scheme of
    TESTBANK_CH01.docx (see module docstring)."""

    def build(self, template_path: str, items: List[OutputItem], output_path: str) -> None:
        docx_template_path = ensure_docx(template_path)
        document = Document(docx_template_path)
        self._clear_body(document)

        for item in items:
            self._write_item(document, item)

        document.save(output_path)
        logger.info("Saved output document with %d items to %s", len(items), output_path)

    # -- internals ---------------------------------------------------------

    @staticmethod
    def _clear_body(document: Document) -> None:
        body = document.element.body
        for child in list(body):
            if child.tag.endswith("}sectPr"):
                continue
            body.remove(child)

    def _write_item(self, document: Document, item: OutputItem) -> None:
        # Outer 1x1 shell table (no borders — matches the source file).
        outer = document.add_table(rows=1, cols=1)
        outer.autofit = True
        cell = outer.cell(0, 0)
        _clear_default_paragraph(cell)

        if item.decision == "REVISE":
            p = cell.add_paragraph()
            run = p.add_run("[REVISE]")
            run.bold = True
            run.font.color.rgb = _COLOR_REVISE_TAG
        elif item.decision == "REMOVE":
            p = cell.add_paragraph()
            run = p.add_run("[STATUS: REMOVED]")
            run.bold = True
            run.font.color.rgb = _COLOR_DEL

        # Stem — tracked-change runs rendered inline (struck-through for REMOVE).
        stem_p = cell.add_paragraph()
        if item.decision == "REMOVE":
            render_fully_struck(stem_p, item.stem)
        else:
            render_tracked_runs(stem_p, item.stem)

        # Nested table 1 — options (skipped for essay items).
        if not item.is_essay:
            self._write_options_table(cell, item)

        cell.add_paragraph("")  # spacer

        # Nested table 2 — metadata.
        self._write_metadata_table(cell, item)

        cell.add_paragraph("")  # spacer

        # AI Review block.
        if item.ai_review_text.strip():
            self._write_ai_review(cell, item.ai_review_text)

        cell.add_paragraph("")  # trailing spacer

    def _write_options_table(self, cell, item: OutputItem) -> None:
        table = cell.add_table(rows=0, cols=3)
        _try_apply_style(table, _NESTED_TABLE_STYLE)
        _set_column_widths(table, [Inches(0.05), Inches(0.3), Inches(6.1)])  # narrow, narrow, wide
        # Only write rows for options the item actually has — e.g. True/False
        # items only have 'a'/'b'. Looping over the full OPTION_LETTERS
        # constant unconditionally is what was producing phantom empty
        # c./d. rows (and throwing off column alignment) on 2-option items.
        present_letters = [l for l in OPTION_LETTERS if (item.options.get(l) or "").strip()]
        for letter in present_letters:
            if letter not in item.options:
                continue
            row = table.add_row()
            blank_cell, letter_cell, text_cell = row.cells
            _clear_default_paragraph(blank_cell)
            _clear_default_paragraph(letter_cell)
            _clear_default_paragraph(text_cell)
            blank_cell.add_paragraph(NBSP)
            letter_p = letter_cell.add_paragraph()
            text_p = text_cell.add_paragraph()
            if item.decision == "REMOVE":
                render_fully_struck(letter_p, f"{letter}.{NBSP}")
                render_fully_struck(text_p, item.options.get(letter, ""))
            else:
                letter_p.add_run(f"{letter}.{NBSP}")
                render_tracked_runs(text_p, item.options.get(letter, ""))

    def _write_metadata_table(self, cell, item: OutputItem) -> None:
        table = cell.add_table(rows=0, cols=2)
        _try_apply_style(table, _NESTED_TABLE_STYLE)

        rows: List[Tuple[str, str]] = [("answer", item.answer)]
        rows += [(f, getattr(item, f)) for f in METADATA_FIELDS if f != "answer"]
        rows.append(("other", item.other_html))

        for field_name, value in rows:
            row = table.add_row()
            label_cell, value_cell = row.cells
            _clear_default_paragraph(label_cell)
            _clear_default_paragraph(value_cell)

            label_p = label_cell.add_paragraph()
            label_run = label_p.add_run(METADATA_LABELS[field_name])
            label_run.italic = True

            value_text = value or ""
            lines = value_text.split("\n") if value_text else [""]
            for line in lines:
                line_p = value_cell.add_paragraph()
                render_tracked_runs(line_p, line)

    def _write_ai_review(self, cell, ai_review_text: str) -> None:
        heading_p = cell.add_paragraph()
        heading_p.paragraph_format.space_before = Pt(0)
        heading_p.paragraph_format.space_after = Pt(0)
        heading_run = heading_p.add_run("[AI REVIEW]")
        heading_run.bold = True
        heading_run.font.color.rgb = _COLOR_REVIEW

        for line in ai_review_text.split("\n"):
            if not line.strip():
                continue
            paragraph = cell.add_paragraph()
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)

            distractor_match = _DISTRACTOR_LINE.match(line.strip())
            label_match = _match_review_label(line)

            if label_match:
                label, rest = label_match
                rest = rest.strip()
                label_run = paragraph.add_run(label + (" " if rest else ""))
                label_run.bold = True
                label_run.font.color.rgb = _COLOR_REVIEW
                if rest:
                    value_run = paragraph.add_run(rest)
                    value_run.bold = False
                    value_run.font.color.rgb = _COLOR_REVIEW
            elif distractor_match:
                letter, rest = distractor_match.groups()
                label_run = paragraph.add_run(f"   {letter.lower()}. ")
                label_run.bold = True
                label_run.font.color.rgb = _COLOR_REVIEW
                value_run = paragraph.add_run(rest)
                value_run.bold = False
                value_run.font.color.rgb = _COLOR_REVIEW
            else:
                run = paragraph.add_run(line)
                run.bold = False
                run.font.color.rgb = _COLOR_REVIEW


# =====================================================================
# =====================================================================
#  SECTION 3 — GEMINI ENGINE
# =====================================================================
# =====================================================================

# =====================================================================
# Gemini API client
# =====================================================================

@dataclass
class InvokeResult:
    text: str
    input_tokens: int
    output_tokens: int

    @property
    def total_tokens(self) -> int:
        return self.input_tokens + self.output_tokens


class GeminiInvocationError(RuntimeError):
    """Raised for any failure invoking Gemini, with a user-facing message."""


class GeminiClient:
    """Thin wrapper around the Gemini API `generate_content` call.

    Supports two auth paths, selected by the GOOGLE_GENAI_USE_VERTEXAI
    environment variable:

      - Vertex AI (service account) — GOOGLE_GENAI_USE_VERTEXAI=true, plus
        GOOGLE_CLOUD_PROJECT and GOOGLE_CLOUD_LOCATION. Credentials come from
        Application Default Credentials, i.e. the GOOGLE_APPLICATION_CREDENTIALS
        env var pointing at a service-account JSON key (e.g. rs-translation.json).
        No API key is needed on this path.

      - Gemini Developer API (API key) — the default when
        GOOGLE_GENAI_USE_VERTEXAI is unset/false. Requires GEMINI_API_KEY.
    """

    def __init__(self, model_id: Optional[str] = None, temperature: Optional[float] = None):
        self.use_vertexai = os.getenv("GOOGLE_GENAI_USE_VERTEXAI", "").lower() in ("1", "true", "yes")
        self.project = os.getenv("GOOGLE_CLOUD_PROJECT")
        self.location = os.getenv("GOOGLE_CLOUD_LOCATION", "global")

        if self.use_vertexai:
            if not self.project:
                raise ValueError(
                    "GOOGLE_GENAI_USE_VERTEXAI is set but GOOGLE_CLOUD_PROJECT is missing. "
                    "Set it to your GCP project id (e.g. 'rs-translations')."
                )
            # Application Default Credentials picks up GOOGLE_APPLICATION_CREDENTIALS
            # automatically — no api_key needed on this path.
            self.api_key = None
            self.client = genai.Client(vertexai=True, project=self.project, location=self.location)
        else:
            self.api_key = os.getenv("GEMINI_API_KEY")
            if not self.api_key:
                raise ValueError(
                    "Gemini API key not found. Set GEMINI_API_KEY in your .env file or environment, "
                    "or set GOOGLE_GENAI_USE_VERTEXAI=true to use the service-account/Vertex AI path instead."
                )
            self.client = genai.Client(api_key=self.api_key)

        self.model_id = model_id or os.getenv("GEMINI_MODEL_ID", "gemini-3.5-flash")
        self.temperature = temperature if temperature is not None else float(os.getenv("GEMINI_TEMPERATURE", "0.1"))
        self.max_output_tokens = int(os.getenv("MAX_OUTPUT_TOKENS", "65536"))
        self.rate_limit_max_retries = int(os.getenv("GEMINI_RATE_LIMIT_MAX_RETRIES", "5"))
        self.rate_limit_base_delay = float(os.getenv("GEMINI_RATE_LIMIT_BASE_DELAY_S", "20"))

        logger.info(
            "Initialized Gemini client using model %s via %s",
            self.model_id, "Vertex AI" if self.use_vertexai else "API key",
        )

    @staticmethod
    def _is_rate_limit_error(exc: Exception) -> bool:
        """True for 429 / RESOURCE_EXHAUSTED errors from the google-genai SDK.
        Checked via the status code when the SDK exposes one (ClientError.code),
        falling back to matching the error text for other exception shapes."""
        code = getattr(exc, "code", None)
        if code == 429:
            return True
        text = str(exc)
        return "429" in text or "RESOURCE_EXHAUSTED" in text

    def invoke(self, prompt: str, max_output_tokens: Optional[int] = None) -> InvokeResult:
        if max_output_tokens is None:
            max_output_tokens = self.max_output_tokens

        config = types.GenerateContentConfig(
            temperature=self.temperature,
            max_output_tokens=max_output_tokens,
        )

        logger.info("Sending prompt to Gemini (%d chars, max_output_tokens=%d)", len(prompt), max_output_tokens)
        start = time.time()

        for attempt in range(1, self.rate_limit_max_retries + 1):
            try:
                response = self.client.models.generate_content(
                    model=self.model_id,
                    contents=prompt,
                    config=config,
                )
                break
            except Exception as exc:
                if self._is_rate_limit_error(exc) and attempt < self.rate_limit_max_retries:
                    wait_s = self.rate_limit_base_delay * (2 ** (attempt - 1))
                    logger.warning(
                        "Rate limited by Gemini (attempt %d/%d) — waiting %.0fs before retry: %s",
                        attempt, self.rate_limit_max_retries, wait_s, exc,
                    )
                    time.sleep(wait_s)
                    continue
                raise GeminiInvocationError(f"Gemini API error: {exc}") from exc

        elapsed = time.time() - start
        logger.info("Received Gemini response after %.2fs", elapsed)

        text = response.text
        if not text:
            raise GeminiInvocationError(
                f"Empty response from Gemini. Finish reason: {getattr(response.candidates[0], 'finish_reason', 'unknown') if response.candidates else 'no candidates'}"
            )

        usage = response.usage_metadata
        input_tokens = int(usage.prompt_token_count or 0) if usage else 0
        output_tokens = int(usage.candidates_token_count or 0) if usage else 0
        logger.info("Token usage — input: %d, output: %d, total: %d", input_tokens, output_tokens, input_tokens + output_tokens)

        return InvokeResult(text=text, input_tokens=input_tokens, output_tokens=output_tokens)


# =====================================================================
# Structured Prompt Builder
# =====================================================================

_METADATA_TAGS = {
    "answer": "ANSWER",
    "points": "POINTS",
    "difficulty": "DIFFICULTY",
    "references": "REFERENCES",
    "learning_objectives": "LEARNING_OBJECTIVES",
    "accrediting_standards": "ACCREDITING_STANDARDS",
    "keywords": "KEYWORDS",
}


def _cdata(raw: str) -> str:
    safe = raw.replace("]]>", "]]]]><![CDATA[>")
    return f"<![CDATA[{safe}]]>"


class PromptBuilder:
    """Builds the structured Gemini prompt from a prompt template, the
    extracted textbook chapter text, and a list of TestBankItems."""

    def serialize_item(self, item: TestBankItem) -> str:
        revise_attr = "true" if item.has_revise_marker else "false"
        lines = [f'<ITEM n="{item.number}" revise_marker="{revise_attr}">']
        lines.append(f"<STEM>{escape(item.stem)}</STEM>")

        if item.is_essay:
            lines.append(f"<ANSWER>{escape(item.answer)}</ANSWER>")
        else:
            for letter in OPTION_LETTERS:
                text = item.options.get(letter, "")
                lines.append(f"<{letter.upper()}>{escape(text)}</{letter.upper()}>")
            lines.append(f"<ANSWER>{escape(item.answer)}</ANSWER>")

        for field_name in METADATA_FIELDS:
            if field_name == "answer":
                continue
            tag = _METADATA_TAGS[field_name]
            value = getattr(item, field_name)
            lines.append(f"<{tag}>{escape(value)}</{tag}>")

        lines.append(f"<OTHER>{_cdata(item.other_html)}</OTHER>")
        lines.append("</ITEM>")
        return "\n".join(lines)

    def build_items_block(self, items: List[TestBankItem]) -> str:
        return "\n\n".join(self.serialize_item(item) for item in items)

    def build_full_prompt(self, prompt_template: str, structured_textbook_text: str, items: List[TestBankItem]) -> str:
        items_xml = self.build_items_block(items)
        return (
            f"{prompt_template}\n\n"
            f"<NEW_TEXTBOOK_CHAPTER>\n{structured_textbook_text}\n</NEW_TEXTBOOK_CHAPTER>\n\n"
            f"<PREVIOUS_EDITION_TEST_BANK>\n{items_xml}\n</PREVIOUS_EDITION_TEST_BANK>\n"
        )

    @staticmethod
    def chunk_items(items: List[TestBankItem], chunk_size: int) -> List[List[TestBankItem]]:
        return [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]

    @staticmethod
    def estimate_output_tokens(items: List[TestBankItem]) -> int:
        total_chars = 0
        for item in items:
            field_chars = len(item.stem)
            for v in item.options.values():
                field_chars += len(v)
            field_chars += sum(len(getattr(item, f)) for f in METADATA_FIELDS)
            field_chars += len(item.other_html)
            review_overhead = 3600 if item.is_essay else 2400
            total_chars += field_chars + review_overhead
        return total_chars // 4

    def adaptive_chunk_items(self, items: List[TestBankItem], chunk_size: int, max_output_tokens: int) -> List[List[TestBankItem]]:
        raw_chunks = self.chunk_items(items, chunk_size)
        final_chunks: List[List[TestBankItem]] = []
        for chunk in raw_chunks:
            estimated = self.estimate_output_tokens(chunk)
            if estimated <= max_output_tokens:
                final_chunks.append(chunk)
            else:
                sub_size = max(1, chunk_size // 2)
                while sub_size > 1:
                    test_chunk = chunk[:sub_size]
                    if self.estimate_output_tokens(test_chunk) <= max_output_tokens:
                        break
                    sub_size -= 1
                logger.info(
                    "Adaptive chunker: chunk of %d items (est. %d output tokens) split to sub-batches of %d",
                    len(chunk), estimated, sub_size,
                )
                final_chunks.extend(self.chunk_items(chunk, sub_size))
        return final_chunks


# =====================================================================
# Gemini Response Parser
# =====================================================================

_OUTPUT_ITEM_PATTERN = re.compile(
    r'<OUTPUT_ITEM\s+n="(\d+)"\s+decision="([A-Z_]+)"\s*>(.*?)</OUTPUT_ITEM>', re.DOTALL,
)
_CDATA_OTHER_PATTERN = re.compile(r"<OTHER>\s*<!\[CDATA\[(.*?)\]\]>\s*</OTHER>", re.DOTALL)
_PLAIN_OTHER_PATTERN = re.compile(r"<OTHER>(.*?)</OTHER>", re.DOTALL)


def _extract_tag(block: str, tag: str) -> Optional[str]:
    match = re.search(rf"<{tag}>(.*?)</{tag}>", block, re.DOTALL)
    if match is None:
        return None
    # NOTE: deliberately do NOT html.unescape here before <DEL>/<INS> are
    # rendered — unescape is safe since those tags aren't HTML entities,
    # but keep it explicit that this string may still contain them.
    return html.unescape(match.group(1)).strip()


def _extract_other(block: str) -> str:
    match = _CDATA_OTHER_PATTERN.search(block)
    if match:
        return match.group(1).replace("]]]]><![CDATA[>", "]]>").strip()
    match = _PLAIN_OTHER_PATTERN.search(block)
    if match:
        return html.unescape(match.group(1)).strip()
    return ""


class ResponseParser:
    """Parses Gemini's raw response text into a list of OutputItem objects,
    aligned one-to-one with the TestBankItems that were sent. Never drops
    an item — unparseable blocks fall back to the original item data with
    decision=PARSE_ERROR."""

    def parse(self, response_text: str, original_items: List[TestBankItem]) -> List[OutputItem]:
        blocks_by_number: Dict[int, tuple] = {}
        for match in _OUTPUT_ITEM_PATTERN.finditer(response_text):
            number = int(match.group(1))
            blocks_by_number[number] = (match.group(2), match.group(3))

        results: List[OutputItem] = []
        mismatches = []
        for original in original_items:
            block = blocks_by_number.get(original.number)
            if block is None:
                mismatches.append(original.number)
                results.append(OutputItem.from_test_bank_item(
                    original, decision=DECISION_PARSE_ERROR,
                    ai_review_text="[PARSE_ERROR] No matching <OUTPUT_ITEM> found in Gemini's response.",
                ))
                continue
            try:
                results.append(self._parse_block(original, block[0], block[1]))
            except Exception as exc:  # noqa: BLE001 - never let one bad item crash the run
                logger.warning("Failed to parse OUTPUT_ITEM %s: %s", original.number, exc)
                mismatches.append(original.number)
                results.append(OutputItem.from_test_bank_item(
                    original, decision=DECISION_PARSE_ERROR, ai_review_text=f"[PARSE_ERROR] {exc}",
                ))

        if len(results) != len(original_items):
            logger.error("Response parser output count (%d) != input count (%d)", len(results), len(original_items))
        if mismatches:
            logger.warning("Items requiring PARSE_ERROR fallback: %s", mismatches)

        return results

    def _parse_block(self, original: TestBankItem, decision: str, body: str) -> OutputItem:
        if decision not in VALID_DECISIONS:
            decision = DECISION_PARSE_ERROR

        stem = _extract_tag(body, "STEM") or original.stem

        options: Dict[str, str] = {}
        if not original.is_essay:
            for letter in OPTION_LETTERS:
                value = _extract_tag(body, letter.upper())
                options[letter] = value if value is not None else original.options.get(letter, "")

        answer = _extract_tag(body, "ANSWER")
        if answer is None:
            answer = original.answer

        metadata_kwargs = {}
        for field_name, tag in _METADATA_TAGS.items():
            if field_name == "answer":
                continue
            value = _extract_tag(body, tag)
            metadata_kwargs[field_name] = value if value is not None else getattr(original, field_name)

        other_html = _extract_other(body)
        if not other_html:
            other_html = original.other_html

        ai_review_text = _extract_tag(body, "AI_REVIEW") or ""

        return OutputItem(
            number=original.number, decision=decision, stem=stem, options=options, answer=answer,
            other_html=other_html, ai_review_text=ai_review_text, is_essay=original.is_essay,
            **metadata_kwargs,
        )


# =====================================================================
# Multi-Chunk Processing Pipeline (Anti-Truncation)
# =====================================================================

ProgressCallback = Optional[Callable[[int, int, ChunkStats], None]]


class GeminiChunkedProcessor:
    """Orchestrates chunked Gemini calls for a full test bank, with retry
    and gap validation."""

    def __init__(
        self,
        gemini_client: GeminiClient,
        prompt_builder: Optional[PromptBuilder] = None,
        response_parser: Optional[ResponseParser] = None,
        chunk_size: int = 10,
        max_retries: int = 2,
        retry_chunk_size: int = 5,
    ):
        self.gemini_client = gemini_client
        self.prompt_builder = prompt_builder or PromptBuilder()
        self.response_parser = response_parser or ResponseParser()
        self.chunk_size = chunk_size
        self.max_retries = max_retries
        self.retry_chunk_size = retry_chunk_size

    def process(self, prompt_template: str, textbook_text: str, items: List[TestBankItem], on_chunk_done: ProgressCallback = None) -> List[OutputItem]:
        max_output_tokens = self.gemini_client.max_output_tokens
        chunks = self.prompt_builder.adaptive_chunk_items(items, self.chunk_size, max_output_tokens)
        total_chunks = len(chunks)
        merged: List[OutputItem] = []

        for index, chunk in enumerate(chunks, start=1):
            chunk_stats, chunk_results = self._process_chunk(prompt_template, textbook_text, chunk, chunk_index=index)
            merged.extend(chunk_results)
            logger.info(
                "Chunk %d/%d: sent %d, returned %d (%.1fs)%s",
                index, total_chunks, chunk_stats.items_sent, chunk_stats.items_returned,
                chunk_stats.elapsed_seconds, " [retried]" if chunk_stats.retried else "",
            )
            if on_chunk_done:
                on_chunk_done(index, total_chunks, chunk_stats)

        merged.sort(key=lambda output_item: output_item.number)
        self._validate_no_gaps(items, merged)
        return merged

    def _process_chunk(self, prompt_template, textbook_text, chunk, chunk_index):
        start = time.time()
        results, input_tok, output_tok = self._call_and_parse(prompt_template, textbook_text, chunk)
        retried = False

        if len(results) < len(chunk):
            logger.warning("Chunk %d returned %d/%d items; retrying at reduced batch size", chunk_index, len(results), len(chunk))
            retried = True
            results, retry_in, retry_out = self._retry_with_smaller_batches(prompt_template, textbook_text, chunk)
            input_tok += retry_in
            output_tok += retry_out

        elapsed = time.time() - start
        stats = ChunkStats(
            chunk_index=chunk_index, items_sent=len(chunk), items_returned=len(results),
            elapsed_seconds=elapsed, retried=retried, input_tokens=input_tok, output_tokens=output_tok,
        )
        return stats, results

    def _retry_with_smaller_batches(self, prompt_template, textbook_text, chunk):
        sub_batches = self.prompt_builder.chunk_items(chunk, self.retry_chunk_size)
        results: List[OutputItem] = []
        total_input = total_output = 0
        for sub_batch in sub_batches:
            attempt_results, in_tok, out_tok = self._call_and_parse(prompt_template, textbook_text, sub_batch)
            total_input += in_tok
            total_output += out_tok
            if len(attempt_results) < len(sub_batch):
                logger.error("Sub-batch retry still short (%d/%d); falling back to PARSE_ERROR for missing items", len(attempt_results), len(sub_batch))
            results.extend(attempt_results)
        return results, total_input, total_output

    def _call_and_parse(self, prompt_template, textbook_text, items):
        prompt = self.prompt_builder.build_full_prompt(prompt_template, textbook_text, items)
        result = self.gemini_client.invoke(prompt)
        output_items = self.response_parser.parse(result.text, items)
        return output_items, result.input_tokens, result.output_tokens

    @staticmethod
    def _validate_no_gaps(original_items: List[TestBankItem], merged: List[OutputItem]) -> None:
        original_numbers = {item.number for item in original_items}
        merged_numbers = {item.number for item in merged}
        missing = original_numbers - merged_numbers
        if missing:
            logger.error("Merged output is missing item numbers: %s", sorted(missing))
        extra = merged_numbers - original_numbers
        if extra:
            logger.warning("Merged output has unexpected extra item numbers: %s", sorted(extra))


# =====================================================================
# Content Quality Rules (Cengage Guidelines Enforcement)
# =====================================================================

_KEYWORDS_LEVEL = re.compile(r"Bloom's:\s*(Remember|Understand|Apply|Analyze|Evaluate|Create)", re.IGNORECASE)
_NEGATIVE_STEM_PATTERN = re.compile(r"which of the following is not|all of the following.*?except", re.IGNORECASE)
_REFER_TO_OPTION_PATTERN = re.compile(r"refer to (?:option\s*)?([a-dA-D])\b")
_TRACKED_TAG_STRIP = re.compile(r"</?(?:DEL|INS)>", re.IGNORECASE)
_QUOTED_PHRASE = re.compile(r'"([^"]{4,})"')
_NO_CHANGE_PHRASES = re.compile(r"no (revision|change) needed|unchanged|already match|identical|no substantive", re.IGNORECASE)
_DEL_TAG_PATTERN = re.compile(r"<DEL>(.*?)</DEL>", re.DOTALL | re.IGNORECASE)
_TRACKED_TAG_STRIP_WITH_CONTENT = re.compile(r"<(?:DEL|INS)>.*?</(?:DEL|INS)>", re.DOTALL | re.IGNORECASE)
_PROPER_NOUN_PATTERN = re.compile(r"\b[A-Z][a-zA-Z]{2,}(?:'s)?\b")
_SCENARIO_CUE_PATTERN = re.compile(r"\b(he|she|his|her|they|their|recently|is looking|wants to|decides to|"r"trying to|planning to|posted|search(?:ing)? for)\b",re.IGNORECASE)

@dataclass
class BloomsOverride:
    number: int
    blooms_level: str


@dataclass
class ValidationReport:
    flags_by_item: Dict[int, List[str]] = field(default_factory=dict)

    def add(self, number: int, flag: str) -> None:
        self.flags_by_item.setdefault(number, []).append(flag)

    def summary(self) -> str:
        if not self.flags_by_item:
            return "No items flagged."
        lines = [f"{len(self.flags_by_item)} item(s) flagged:"]
        for number, flags in sorted(self.flags_by_item.items()):
            lines.append(f"  Q{number}: {', '.join(flags)}")
        return "\n".join(lines)


class ContentValidator:
    """Applies Cengage content-quality rules to OutputItems before the docx is built."""

    def __init__(self, blooms_overrides: List[BloomsOverride] = None):
        self.blooms_overrides = {ov.number: ov.blooms_level for ov in (blooms_overrides or [])}
        self._rules: List[Callable[[OutputItem, ValidationReport], None]] = [
            self._check_negative_stem,
            self._check_blooms_override,
            self._check_distractor_length,
            self._check_refer_to_pointer,
            self._check_incomplete_entity_replacement,
            self._check_wording_rule_self_contradiction,
            self._check_named_scenario_blooms_floor,
        ]

    def validate(self, items: List[OutputItem]) -> ValidationReport:
        report = ValidationReport()
        for item in items:
            for rule in self._rules:
                rule(item, report)
        return report

    @staticmethod
    def _plain(text: str) -> str:
        """Strip <DEL>/<INS> tags (keep their inner text) before running
        text-content rules, so tracked-change markup doesn't confuse
        pattern matches like the negative-stem check."""
        return _TRACKED_TAG_STRIP.sub("", text or "")

    def _check_negative_stem(self, item: OutputItem, report: ValidationReport) -> None:
        if _NEGATIVE_STEM_PATTERN.search(self._plain(item.stem)):
            self._append_flag(item, report, "[FLAG: negative stem — convert to positive]")

    def _check_blooms_override(self, item: OutputItem, report: ValidationReport) -> None:
        expected = self.blooms_overrides.get(item.number)
        if not expected:
            return
        if expected.lower() not in (item.keywords or "").lower():
            item.keywords = f"Bloom's: {expected}"
            report.add(item.number, f"[OVERRIDE: Bloom's level forced to {expected}]")

    def _check_distractor_length(self, item: OutputItem, report: ValidationReport) -> None:
        lengths = [len(self._plain(item.options.get(letter) or "").split()) for letter in "abcd"]
        lengths = [length for length in lengths if length > 0]
        if len(lengths) < 2:
            return
        if max(lengths) > 3 * max(min(lengths), 1):
            self._append_flag(item, report, "[FLAG: distractor length inconsistency]")

    def _check_incomplete_entity_replacement(self, item: OutputItem, report: ValidationReport) -> None:
        """A REVISE item that swaps a named entity must catch every occurrence
        of that name — not just the first one. The model is asked to self-verify
        this via the Completeness Check in the prompt, but that self-report has
        proven unreliable in practice, so verify it mechanically here instead."""
        if item.decision != "REVISE":
            return

        full_text = item.stem + " " + " ".join(item.options.values())

        old_names = set()
        for span in _DEL_TAG_PATTERN.findall(full_text):
            for tok in _PROPER_NOUN_PATTERN.findall(span):
                old_names.add(tok.rstrip("'s"))

        if not old_names:
            return

        # Remove every <DEL>...</DEL> and <INS>...</INS> span (tag + content) —
        # whatever text remains is genuinely untouched, unrevised plain text.
        remaining_plain_text = _TRACKED_TAG_STRIP_WITH_CONTENT.sub("", full_text)

        leftover = [
            name for name in old_names
            if re.search(r"\b" + re.escape(name) + r"(?:'s)?\b", remaining_plain_text)
        ]

        if leftover:
            self._append_flag(
                item, report,
                f"[FLAG: incomplete entity replacement — '{', '.join(sorted(leftover))}' "
                f"still appears untracked outside <DEL>/<INS> tags — some occurrence wasn't revised]"
            )

    def _check_wording_rule_self_contradiction(self, item: OutputItem, report: ValidationReport) -> None:
        """If the model's own Wording-vs-Substance line quotes two different
        phrases but still claims nothing changed, that's a self-contradiction —
        catch it mechanically instead of trusting the model's own verdict."""
        match = re.search(r"Wording-vs-Substance Rule:(.*?)(?:\n[A-Z][\w '\-]*:|\Z)", item.ai_review_text or "", re.DOTALL)
        if not match:
            return
        line = match.group(1)
        quotes = _QUOTED_PHRASE.findall(line)
        if len(quotes) >= 2 and quotes[0].strip().lower() != quotes[1].strip().lower():
            if _NO_CHANGE_PHRASES.search(line):
                self._append_flag(item, report,
                    "[FLAG: Wording-vs-Substance Rule quotes two different phrases but claims no change — needs human review]")

    def _check_named_scenario_blooms_floor(self, item: OutputItem, report: ValidationReport) -> None:
        """A stem built around a constructed personal scenario (named people,
        narrative cues like 'recently posted', 'is looking for') and asking
        the reader to apply a concept to it should be Apply or higher.
        A stem that simply states a fact about named companies/entities
        (e.g. a ranking) is legitimate recall and should NOT be flagged —
        that was the source of false positives on items like the Fortune
        ranking item. REMOVE items are skipped entirely since their
        KEYWORDS field is moot."""
        if item.decision == "REMOVE":
            return

        level_match = _KEYWORDS_LEVEL.search(item.keywords or "")
        if not level_match or level_match.group(1).lower() not in ("remember", "understand"):
            return

        plain_stem = self._plain(item.stem)
        has_named_entity = bool(_PROPER_NOUN_PATTERN.search(plain_stem))
        has_scenario_cue = bool(_SCENARIO_CUE_PATTERN.search(plain_stem))
        if not (has_named_entity and has_scenario_cue):
            return

        review = item.ai_review_text or ""
        justification_present = bool(re.search(
            r"does not (?:actually )?require applying|just restat(?:es|ing) a definition|no new (?:case|scenario)",
            review, re.IGNORECASE
        ))
        if justification_present:
            return

        self._append_flag(
            item, report,
            f"[FLAG: stem is a constructed personal scenario but KEYWORDS is Bloom's: {level_match.group(1)} "
            f"with no stated justification for staying below Apply — verify Bloom's level]"
        )

    def _check_refer_to_pointer(self, item: OutputItem, report: ValidationReport) -> None:
        match = _REFER_TO_OPTION_PATTERN.search(item.ai_review_text or "")
        if not match:
            return
        referenced = match.group(1).lower()
        if item.answer and referenced != item.answer.lower():
            self._append_flag(item, report, f"[FLAG: 'refer to {referenced}' does not match ANSWER {item.answer}]")

    @staticmethod
    def _append_flag(item: OutputItem, report: ValidationReport, flag: str) -> None:
        item.flags.append(flag)
        item.ai_review_text = (item.ai_review_text.rstrip() + f"\n{flag}").strip()
        report.add(item.number, flag)


# =====================================================================
# =====================================================================
#  SECTION 4 — PIPELINE ORCHESTRATOR
# =====================================================================
# =====================================================================

# logs/ lives one level above src/, add it to the path so we can import from it
_LOGS_DIR = Path(__file__).parent.parent / "logs"
sys.path.insert(0, str(_LOGS_DIR))      # noqa: E402
from output_logger import append_output_log                    # noqa: E402
from calculate_logs import calculate as calculate_cost_log     # noqa: E402
from cost_logger import CostAccumulator, append_to_excel
import cost_logger

print("COST LOGGER LOADED FROM:", cost_logger.__file__)
print("RunCostRecord fields:", cost_logger.RunCostRecord.__dataclass_fields__.keys())


logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

PipelineProgressCallback = Optional[Callable[[str, int, int], None]]
# Callback signature: (status_message, chunks_done, total_chunks)

DEFAULT_BLOOMS_OVERRIDES = [
    # Verified directly against TESTBANK_CH01.docx (the reference output),
    # not the original problem statement's example values, which didn't
    # match this file: Q9 -> "Understand" and Q19 -> "Remember".
    BloomsOverride(number=9, blooms_level="Understand"),
    BloomsOverride(number=19, blooms_level="Remember"),
]

# Both log files live in the logs/ folder so they persist across runs
COST_LOG_FILENAME   = "logs/cost_log.xlsx"
OUTPUT_LOG_FILENAME = "logs/output_log.xlsx"


@dataclass
class PipelineResult:
    output_docx_path: str
    raw_response_debug_path: str
    items: List[OutputItem]
    total_items: int
    revise_count: int
    remove_count: int
    parse_error_count: int
    cost_log_path: str = ""      # path to logs/cost_log.xlsx
    output_log_path: str = ""    # path to logs/output_log.xlsx
    calculate_log_path: str = "" # path to logs/cost_log.xlsx (Summary sheet refreshed)


class TestBankPipeline:
    """Main pipeline for the test bank audit. Construct once (it owns the
    Gemini client) and call `process_test_bank` per job."""

    def __init__(
        self,
        base_path: Optional[str] = None,
        model_id: Optional[str] = None,
        temperature: Optional[float] = None,
        chunk_size: Optional[int] = None,
        max_retries: Optional[int] = None,
    ):
        self.gemini_client = GeminiClient(model_id=model_id, temperature=temperature)
        self.test_bank_parser = TestBankParser()
        self.chapter_extractor = ChapterExtractor()
        self.prompt_builder = PromptBuilder()
        self.response_parser = ResponseParser()
        self.content_validator = ContentValidator(blooms_overrides=DEFAULT_BLOOMS_OVERRIDES)
        self.docx_builder = DocxBuilder()

        chunk_size = chunk_size or int(os.getenv("CHUNK_SIZE", "10"))
        max_retries = max_retries if max_retries is not None else int(os.getenv("MAX_RETRIES", "2"))
        retry_chunk_size = int(os.getenv("RETRY_CHUNK_SIZE", "5"))
        self.chunked_processor = GeminiChunkedProcessor(
            gemini_client=self.gemini_client,
            prompt_builder=self.prompt_builder,
            response_parser=self.response_parser,
            chunk_size=chunk_size,
            max_retries=max_retries,
            retry_chunk_size=retry_chunk_size,
        )

        self.base_path = Path(base_path) if base_path else Path(__file__).parent.parent.resolve()

    def load_prompt(self, prompt_file: str = "prompts/testbank.txt") -> str:
        prompt_path = self.base_path / prompt_file
        with open(prompt_path, "r", encoding="utf-8") as handle:
            return handle.read()

    def process_test_bank(
        self,
        textbook_chapter_file: str,
        previous_test_bank_file: str,
        prompt_file: str = "prompts/testbank.txt",
        output_file: Optional[str] = None,
        output_dir: str = "temp_uploads",
        on_progress: PipelineProgressCallback = None,
    ) -> PipelineResult:
        """Run the full audit pipeline and return a PipelineResult.

        Args:
            textbook_chapter_file: path to the NEW edition chapter (.doc/.docx/.pdf)
            previous_test_bank_file: path to the PREVIOUS edition test bank (.doc/.docx)
        """

        def notify(message: str, done: int = 0, total: int = 0):
            logger.info(message)
            if on_progress:
                on_progress(message, done, total)

        # ---- Set up cost accumulator for this run -------------------------
        cost_acc = CostAccumulator(
            model_id=self.gemini_client.model_id,
            region="global",
            test_bank_file=previous_test_bank_file,
            textbook_file=textbook_chapter_file,
        )

        notify("Loading review prompt...")
        prompt_template = self.load_prompt(prompt_file)

        notify(f"Parsing test bank items from {Path(previous_test_bank_file).name}...")
        items: List[TestBankItem] = self.test_bank_parser.parse(previous_test_bank_file)
        notify(f"Parsed {len(items)} test bank items.")

        notify(f"Extracting textbook chapter from {Path(textbook_chapter_file).name}...")
        textbook_text = self.chapter_extractor.extract(textbook_chapter_file)

        total_chunks = max(1, -(-len(items) // self.chunked_processor.chunk_size))  # ceil division

        def chunk_progress(chunk_index: int, total: int, stats: ChunkStats):
            cost_acc.add_chunk(stats)
            notify(f"Processed chunk {chunk_index}/{total} ({stats.items_returned} items)...", chunk_index, total)

        notify(f"Sending {len(items)} items to Gemini in {total_chunks} chunk(s)...", 0, total_chunks)
        output_items = self.chunked_processor.process(
            prompt_template, textbook_text, items, on_chunk_done=chunk_progress
        )

        notify("Applying content quality rules...")
        report = self.content_validator.validate(output_items)
        if report.flags_by_item:
            logger.info("Content validator: %s", report.summary())

        notify("Building output document...")
        if output_file is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"{Path(previous_test_bank_file).stem}_AUDITED_{timestamp}.docx"

        output_dir_path = self.base_path / output_dir
        output_dir_path.mkdir(parents=True, exist_ok=True)
        output_path = output_dir_path / output_file

        self.docx_builder.build(previous_test_bank_file, output_items, str(output_path))

        debug_path = output_path.with_suffix(".debug.txt")
        with open(debug_path, "w", encoding="utf-8") as handle:
            handle.write(self._debug_dump(output_items, report))

        # ---- Finalise cost record and write Excel log ---------------------
        keep_count  = sum(1 for i in output_items if i.decision == "KEEP")
        revise_count  = sum(1 for i in output_items if i.decision == "REVISE")
        remove_count  = sum(1 for i in output_items if i.decision == "REMOVE")
        parse_err_count = sum(1 for i in output_items if i.decision == "PARSE_ERROR")

        cost_record = cost_acc.finalise(
            total_items=len(output_items),
            keep=keep_count,
            revise=revise_count,
            remove=remove_count,
            parse_errors=parse_err_count,
        )
        cost_log_path = str(self.base_path / COST_LOG_FILENAME)
        append_to_excel(cost_record, cost_log_path)
        logger.info(
            "Cost summary — input: %d tok, output: %d tok, total: $%.4f",
            cost_record.total_input_tokens,
            cost_record.total_output_tokens,
            cost_record.total_cost,
        )

        # ---- Recalculate Summary sheet in the cost log -------------------
        calculate_log_path = ""
        try:
            calculate_log_path = calculate_cost_log(cost_log_path)
            logger.info("Cost log Summary sheet refreshed: %s", calculate_log_path)
        except Exception as exc:
            logger.warning("Could not refresh cost log Summary sheet: %s", exc)

        # ---- Write output decisions log ----------------------------------
        output_log_path = str(self.base_path / OUTPUT_LOG_FILENAME)
        append_output_log(
            run_id=cost_record.run_id,
            timestamp=cost_record.timestamp,
            test_bank_file=previous_test_bank_file,
            output_items=output_items,
            duration_s=cost_record.duration_s,
            log_path=output_log_path,
        )

        notify("Done!")

        return PipelineResult(
            output_docx_path=str(output_path),
            raw_response_debug_path=str(debug_path),
            items=output_items,
            total_items=len(output_items),
            revise_count=revise_count,
            remove_count=remove_count,
            parse_error_count=parse_err_count,
            cost_log_path=cost_log_path,
            output_log_path=output_log_path,
            calculate_log_path=calculate_log_path,
        )

    @staticmethod
    def _debug_dump(items: List[OutputItem], report) -> str:
        lines = [f"Total items: {len(items)}"]
        for decision in ("KEEP", "REVISE", "REMOVE", "PARSE_ERROR"):
            count = sum(1 for item in items if item.decision == decision)
            lines.append(f"{decision}: {count}")
        lines.append("")
        lines.append(report.summary())
        return "\n".join(lines)


def main():
    """CLI entry point — supply your own file paths when running directly.

    Example:
        python testbank_pipeline.py
    Edit the two paths below to point at your actual input files.
    """
    pipeline = TestBankPipeline()
    logger.info("Working directory: %s", pipeline.base_path)
    logger.info("Model: %s", pipeline.gemini_client.model_id)

    # ---- Edit these paths before running from the CLI -------------------
    TEXTBOOK_CHAPTER   = r"path/to/new_edition_chapter.docx"
    PREVIOUS_TEST_BANK = r"path/to/previous_test_bank.docx"
    # ---------------------------------------------------------------------

    result = pipeline.process_test_bank(
        textbook_chapter_file=TEXTBOOK_CHAPTER,
        previous_test_bank_file=PREVIOUS_TEST_BANK,
        prompt_file="prompts/testbank.txt",
    )

    logger.info("Pipeline completed. Output: %s", result.output_docx_path)
    logger.info("Cost log: %s", result.cost_log_path)
    logger.info(
        "Items: %d total, %d revise, %d remove, %d parse errors",
        result.total_items, result.revise_count, result.remove_count, result.parse_error_count,
    )


if __name__ == "__main__":
    main()