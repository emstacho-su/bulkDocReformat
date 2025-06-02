# modernizer/parser.py

from pathlib import Path
from typing import List, Dict, Any
import re
import logging

from docx import Document

# Configure basic logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Ordered top‐level keywords (lowercase)
TOP_LEVEL_SEQUENCE = [
    "purpose",
    "scope",
    "definitions",
    "process owner",
    "procedures",
    "references",
    "records",
    "related documents",
    "revisions",
]

# Regex to strip any leading numeric prefix (e.g. "2.", "3.1", "4.")
NUM_PREFIX_PATTERN = re.compile(r"^[\d\.]+\s*(.*)$")

# Regex to match subclause pattern "x.x" (e.g. "4.1", "5.2", "6.3.1")
SUBCLAUSE_PATTERN = re.compile(r"^\s*\d+\.\d+")

# Regex to match sub‐subclause pattern "x.x.x" (e.g. "4.1.1", "5.2.3")
SUBSUBCLAUSE_PATTERN = re.compile(r"^\s*\d+\.\d+\.\d+")

# Regex to match single‐level prefix "7. Something" but not "7.1"
SINGLE_LEVEL_PATTERN = re.compile(r"^(\d+)\.\s+")

# Case‐insensitive match for "process owner"
PROCESS_OWNER_PATTERN = re.compile(r"process\s+owner", re.IGNORECASE)

# Case‐insensitive match for lines beginning with "process designee"
PROCESS_DESIGNEE_PATTERN = re.compile(r"^process\s+designee", re.IGNORECASE)


def normalize(text: str) -> str:
    return text.strip().lower()


def extract_revision_history(doc: Document) -> Dict[str, Any]:
    """
    If the last element in the document is a table, return {"type": "table", "rows": [...]},
    otherwise return {"type": "written", "rows": []}.
    """
    if doc.tables:
        last_table = doc.tables[-1]
        rows: List[List[str]] = []
        for row in last_table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows.append(cells)
        return {"type": "table", "rows": rows}
    return {"type": "written", "rows": []}


def strip_numeric_prefix(text: str) -> str:
    """
    Remove a leading numeric prefix (e.g. "4.1", "5.", "6.") plus whitespace;
    if none, return the original text.
    """
    m = NUM_PREFIX_PATTERN.match(text.strip())
    if m:
        return m.group(1).strip()
    return text.strip()


def parse_legacy_docx_by_sequence(path: Path) -> Dict[str, Any]:
    """
    Parses a legacy .docx with the following rules:

    1. A paragraph starting with a single‐level numeric prefix (e.g. "7. Policy References",
       "8. Revisions…") → new top‐level node, provided we are not inside a subclause.
    2. Otherwise, if a paragraph’s stripped remainder begins with the next keyword in
       TOP_LEVEL_SEQUENCE, it becomes that top‐level node.
    3. If “definitions” never appears, but a bold “Term: Definition” or a subclause shows up,
       create a synthetic "2. Definitions" top‐level.
    4. Inside “Definitions,” a paragraph containing “:” → new child. A paragraph without “:”
       immediately following is appended onto the previous child’s heading (continuation).
    5. “Process Owner” → next non‐empty paragraph becomes the **single child** under “Process Owner.”
    6. “Process Designee” → new top‐level; all following lines (until next top‐level or “Procedures”)
       get lumped as one child under “Process Designee.”
    7. A line matching “x.x.x …” (three‐level numeric) → a **sub‐subclause** under the active
       `current_sub` (if `current_sub` exists). That sub‐subclause becomes `current_subsub`.
    8. A line matching “x.x …” (two‐level numeric) → a **subclause** under the active top‐level.
        - This resets `current_subsub = None`.
    9. Once inside a subclause (or sub‐subclause), any further bold text → appended as content to
       the deepest active node (i.e. `current_subsub` if exists, else `current_sub`).
   10. Any other bold line (when not capturing Process Owner/Designee) → a new subclause under
       the current top‐level (unless a subclause is open).
   11. Non‐bold lines attach as content under the deepest active node (`current_subsub` → `current_sub` → `current_top`).
   12. Finally, extract revision history via a trailing table, or treat leftover text as a
       written revision block.
    """
    try:
        doc = Document(path)
    except Exception as e:
        logger.error(f"Failed to open {path.name}: {e}")
        raise

    paras = [p for p in doc.paragraphs if p.text.strip()]
    if not paras:
        logger.warning(f"No paragraphs found in {path.name}")
        return {"document_title": "", "sections": [], "revision_history": {"type": "written", "content": []}}

    # (1) Document title = first bold paragraph if present, else first paragraph
    document_title = ""
    for p in paras:
        if any(run.bold for run in p.runs):
            document_title = p.text.strip()
            break
    if not document_title:
        document_title = paras[0].text.strip()
        logger.warning(f"No bold title in {path.name}; using first line")

    logger.info(f"[{path.name}] Document Title: {document_title!r}")

    sections: List[Dict[str, Any]] = []
    current_top: Dict[str, Any] = None
    current_sub: Dict[str, Any] = None
    current_subsub: Dict[str, Any] = None
    next_top_idx = 0  # index into TOP_LEVEL_SEQUENCE

    # Flags for capturing “Process Owner” child
    expect_owner_child = False

    # Flags for capturing “Process Designee” block
    capturing_designee_block = False
    designee_lines: List[str] = []

    def create_top(heading: str, forced_keyword: str = None) -> None:
        """
        Create a new top‐level node with "heading". If forced_keyword is provided,
        advance next_top_idx to that keyword’s index+1. Otherwise, if stripped remainder
        matches TOP_LEVEL_SEQUENCE[next_top_idx], advance by 1.

        Reset current_sub and current_subsub.
        """
        nonlocal sections, current_top, current_sub, current_subsub
        nonlocal next_top_idx, expect_owner_child, capturing_designee_block, designee_lines

        node = {"heading": heading, "content": [], "children": []}
        sections.append(node)
        current_top = node
        current_sub = None
        current_subsub = None
        expect_owner_child = False
        capturing_designee_block = False
        designee_lines = []

        if forced_keyword:
            try:
                idx = TOP_LEVEL_SEQUENCE.index(forced_keyword.lower())
                next_top_idx = idx + 1
            except ValueError:
                pass
        else:
            rem = normalize(strip_numeric_prefix(heading))
            if next_top_idx < len(TOP_LEVEL_SEQUENCE) and rem.startswith(TOP_LEVEL_SEQUENCE[next_top_idx]):
                next_top_idx += 1

    # (2) Iterate through all non‐empty paragraphs
    for p in paras:
        text = p.text.strip()
        if not text:
            continue

        stripped = strip_numeric_prefix(text)
        low = normalize(stripped)
        raw = text.strip()
        is_bold = any(run.bold for run in p.runs)

        # ----- A) If capturing a Designee block, collect lines until a top‐level or subclause appears -----
        if capturing_designee_block:
            # If this paragraph signals “Procedures” (next keyword) or a single‐level numeric prefix
            # (and we’re NOT inside a subclause), flush designee_lines and re‐process:
            if (low.startswith("procedures")
                   and next_top_idx < len(TOP_LEVEL_SEQUENCE)
                   and TOP_LEVEL_SEQUENCE[next_top_idx] == "procedures") \
               or (current_sub is None and SINGLE_LEVEL_PATTERN.match(raw) and not SUBCLAUSE_PATTERN.match(raw)):
                child_heading = "\n".join(designee_lines).strip()
                if child_heading:
                    child = {"heading": child_heading, "content": [], "children": []}
                    current_top["children"].append(child)
                designee_lines = []
                capturing_designee_block = False
                # fall through to re‐process this paragraph as top‐level or subclause
            else:
                designee_lines.append(text)
                continue

        # ----- B) SINGLE‐LEVEL numeric prefix → NEW TOP-LEVEL (only if NOT inside a subclause) -----
        if current_sub is None and SINGLE_LEVEL_PATTERN.match(raw) and not SUBCLAUSE_PATTERN.match(raw):
            create_top(text)
            continue

        # ----- C) TOP-LEVEL keyword in sequence -----
        if next_top_idx < len(TOP_LEVEL_SEQUENCE):
            expected = TOP_LEVEL_SEQUENCE[next_top_idx]
            if low.startswith(expected):
                # If it’s “process designee,” go to block E instead
                if low.startswith("process designee"):
                    pass
                else:
                    create_top(text)
                    if low.startswith("process owner"):
                        expect_owner_child = True
                    continue

        # ----- D) PROCESS OWNER: next non‐empty paragraph is its single child -----
        if expect_owner_child:
            child = {"heading": text, "content": [], "children": []}
            current_top["children"].append(child)
            current_sub = child
            current_subsub = None
            expect_owner_child = False
            continue

        # ----- E) PROCESS DESIGNEE: new top‐level; collect subsequent lines in one block -----
        if low.startswith("process designee"):
            create_top(text)
            capturing_designee_block = True
            designee_lines = []
            continue

        # ----- F) “DEFINITIONS” section: paragraphs with “:” → new child; others append to last child -----
        if current_top and normalize(strip_numeric_prefix(current_top["heading"])).startswith("definitions"):
            # Check if would actually be a new top‐level (“Procedures” or single‐level numeric):
            is_new_top = (current_sub is None and SINGLE_LEVEL_PATTERN.match(raw) and not SUBCLAUSE_PATTERN.match(raw)) \
                         or (next_top_idx < len(TOP_LEVEL_SEQUENCE) and low.startswith(TOP_LEVEL_SEQUENCE[next_top_idx]))
            if not is_new_top:
                # If this paragraph CONTAINS a colon → start a BRAND‐NEW definition child
                if ":" in text:
                    child = {"heading": text, "content": [], "children": []}
                    current_top["children"].append(child)
                    current_sub = child
                    current_subsub = None
                else:
                    # continuation of the previous definition: append onto its heading
                    if current_sub:
                        current_sub["heading"] += " " + text
                    else:
                        # If somehow no current_sub, just treat as a new child
                        child = {"heading": text, "content": [], "children": []}
                        current_top["children"].append(child)
                        current_sub = child
                        current_subsub = None
                continue

        # ----- G) SUB-SUBCLAUSE detection "x.x.x" → new node under current_sub (if exists) -----
        if SUBSUBCLAUSE_PATTERN.match(raw) and current_sub:
            child = {"heading": text, "content": [], "children": []}
            current_sub["children"].append(child)
            current_subsub = child
            continue

        # ----- H) SUBCLAUSE detection "x.x" → new node under current_top (resets sub-sub) -----
        if SUBCLAUSE_PATTERN.match(raw):
            if current_top:
                child = {"heading": text, "content": [], "children": []}
                current_top["children"].append(child)
                current_sub = child
                current_subsub = None
            continue

        # ----- I) Any other bold (outside Process Owner/Designee) → new subclause if not inside one -----
        if is_bold:
            # If inside sub‐subclause, append to its content
            if current_subsub:
                current_subsub["content"].append(text)
                continue
            # If inside subclause, append to its content
            if current_sub:
                current_sub["content"].append(text)
                continue
            # Otherwise, create a new subclause under current_top
            if current_top:
                child = {"heading": text, "content": [], "children": []}
                current_top["children"].append(child)
                current_sub = child
                current_subsub = None
            continue

        # ----- J) Non‐bold → attach to deepest active node (sub‐sub → sub → top) -----
        if current_subsub:
            current_subsub["content"].append(text)
        elif current_sub:
            current_sub["content"].append(text)
        elif current_top:
            current_top["content"].append(text)
        else:
            # No parent → skip
            pass

    # (3) Flush any remaining designee_lines
    if capturing_designee_block and current_top and designee_lines:
        child_heading = "\n".join(designee_lines).strip()
        if child_heading:
            child = {"heading": child_heading, "content": [], "children": []}
            current_top["children"].append(child)

    # (4) Extract revision history
    rev_hist = extract_revision_history(doc)
    if rev_hist["type"] == "table":
        if current_subsub:
            current_subsub["content"] = []
        elif current_sub:
            current_sub["content"] = []
        elif current_top:
            current_top["content"] = []
        return {"document_title": document_title, "sections": sections, "revision_history": rev_hist}

    # (5) Otherwise, leftover content is written revision history
    if current_subsub:
        written = {"type": "written", "content": current_subsub["content"]}
        current_subsub["content"] = []
    elif current_sub:
        written = {"type": "written", "content": current_sub["content"]}
        current_sub["content"] = []
    elif current_top:
        written = {"type": "written", "content": current_top["content"]}
        current_top["content"] = []
    else:
        written = {"type": "written", "content": []}

    return {"document_title": document_title, "sections": sections, "revision_history": written}
