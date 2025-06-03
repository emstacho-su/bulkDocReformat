from pathlib import Path
from typing import List, Dict, Any
import re
import logging

from docx import Document

# ---------------------------------------------------------------------
#  Logging
# ---------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------
#  Ordered top‑level keywords (lower‑case, stripped of numeric prefixes)
# ---------------------------------------------------------------------
TOP_LEVEL_SEQUENCE = [
    "purpose",          # handled specially
    "scope",
    "definitions",
    "process owner",
    "procedures",
    "references",
    "related documents",
    "records",
    "policy reference",
    "revisions",
]

# ---------------------------------------------------------------------
#  Regex helpers
# ---------------------------------------------------------------------
NUM_PREFIX_PATTERN    = re.compile(r"^[\d\.]+\s*(.*)$")
SUBCLAUSE_PATTERN     = re.compile(r"^\s*\d+\.\d+")
SUBSUBCLAUSE_PATTERN  = re.compile(r"^\s*\d+\.\d+\.\d+")
SINGLE_LEVEL_PATTERN  = re.compile(r"^(\d+)\.\s+")
PURPOSE_AND_SCOPE_PATTERN = re.compile(r"^purpose\s+and\s+scope", re.IGNORECASE)
PROCESS_DESIGNEE_PATTERN  = re.compile(r"^process\s+designee", re.IGNORECASE)

# ---------------------------------------------------------------------
#  Utility functions
# ---------------------------------------------------------------------
def normalize(text: str) -> str:
    return text.strip().lower()

def strip_numeric_prefix(text: str) -> str:
    m = NUM_PREFIX_PATTERN.match(text.strip())
    return m.group(1).strip() if m else text.strip()

def extract_revision_history(doc: Document) -> Dict[str, Any]:
    """Return {"type":"table",...} if last element is table, else written."""
    if doc.tables:
        rows = [
            [cell.text.strip() for cell in row.cells]
            for row in doc.tables[-1].rows
        ]
        return {"type": "table", "rows": rows}
    return {"type": "written", "content": []}

# ---------------------------------------------------------------------
#  Main parser
# ---------------------------------------------------------------------
def parse_legacy_docx_by_sequence(path: Path) -> Dict[str, Any]:
    """
    Parse a legacy .docx file and return a structure for populate_template.
    """

    # A)  Load Word file
    try:
        doc = Document(path)
    except Exception as e:
        logger.error(f"Failed to open {path.name}: {e}")
        raise

    # B)  Collect non‑blank paragraphs
    paras = [p for p in doc.paragraphs if p.text.strip()]
    if not paras:
        logger.warning(f"No paragraphs found in {path.name}")
        return {
            "document_title": "",
            "purpose_scope_block": "",
            "sections": [],
            "revision_history": {"type": "written", "content": []},
        }

    # C)  Document title = first bold paragraph (fallback to first line)
    document_title = next(
        (p.text.strip() for p in paras if any(r.bold for r in p.runs)), paras[0].text.strip()
    )
    if document_title == paras[0].text.strip():
        logger.warning(f"No bold title in {path.name}; using first line")

    logger.info(f"[{path.name}] Document Title: {document_title!r}")

    # D)  Locate “Purpose and Scope” and “Definitions”
    idx_ps, idx_def = None, None
    for idx, p in enumerate(paras):
        stripped = strip_numeric_prefix(p.text)
        low = normalize(stripped)
        if idx_ps is None and PURPOSE_AND_SCOPE_PATTERN.match(stripped):
            idx_ps = idx
        if idx_def is None and low.startswith("definitions"):
            idx_def = idx
            break

    if idx_ps is None or idx_def is None or idx_def <= idx_ps:
        logger.warning("Couldn't find both 'Purpose and Scope' and 'Definitions'; "
                       "skipping Purpose/Scope extraction.")
        purpose_scope_block = ""
        idx_def = idx_def if idx_def is not None else len(paras)
    else:
        purpose_scope_block = "\n".join(
            p.text.strip() for p in paras[idx_ps + 1 : idx_def] if p.text.strip()
        )

    # E)  Iterate from Definitions onward, building nested section structure
    remaining = paras[idx_def:]
    sections: List[Dict[str, Any]] = []

    current_top = current_sub = current_subsub = None
    next_top_idx = 2  # 'purpose' and 'scope' consumed logically

    expect_owner_child = False
    capturing_designee_block = False
    designee_buffer: List[str] = []

    def create_top(heading: str):
        nonlocal sections, current_top, current_sub, current_subsub
        nonlocal next_top_idx, expect_owner_child, capturing_designee_block, designee_buffer
        node = {"heading": heading, "content": [], "children": []}
        sections.append(node)
        current_top, current_sub, current_subsub = node, None, None
        expect_owner_child = capturing_designee_block = False
        designee_buffer = []

        stripped_head = normalize(strip_numeric_prefix(heading))
        if next_top_idx < len(TOP_LEVEL_SEQUENCE) and stripped_head.startswith(
            TOP_LEVEL_SEQUENCE[next_top_idx]
        ):
            next_top_idx += 1

    for p in remaining:
        text = p.text.strip()
        if not text:
            continue

        stripped = strip_numeric_prefix(text)
        low = normalize(stripped)
        raw = text
        is_bold = any(run.bold for run in p.runs)

        # ------------------------------------------------------------------
        #  Special case: bold "Records" without numeric prefix
        # ------------------------------------------------------------------
        if is_bold and low == "records":
            create_top(text)
            continue

        # ------------------------------------------------------------------
        #  Handle ongoing Process‑Designee multi‑line block
        # ------------------------------------------------------------------
        if capturing_designee_block:
            is_new_top = (
                current_sub is None
                and SINGLE_LEVEL_PATTERN.match(raw)
                and not SUBCLAUSE_PATTERN.match(raw)
            ) or (
                next_top_idx < len(TOP_LEVEL_SEQUENCE)
                and low.startswith(TOP_LEVEL_SEQUENCE[next_top_idx])
            )
            if is_new_top:
                for line in designee_buffer:
                    token = line.strip()
                    if token:
                        current_top["children"].append(
                            {"heading": token, "content": [], "children": []}
                        )
                designee_buffer = []
                capturing_designee_block = False
            else:
                designee_buffer.append(text)
                continue

        # ------------------------------------------------------------------
        #  Single‑level numeric prefix → new top
        # ------------------------------------------------------------------
        if SINGLE_LEVEL_PATTERN.match(raw) and not SUBCLAUSE_PATTERN.match(raw):

            create_top(text)
            if normalize(strip_numeric_prefix(text)).startswith("process owner"):
                expect_owner_child = True
            continue

        # ------------------------------------------------------------------
        #  Keyword in sequence → new top
        # ------------------------------------------------------------------
        if next_top_idx < len(TOP_LEVEL_SEQUENCE):
            expected = TOP_LEVEL_SEQUENCE[next_top_idx]
            if low.startswith(expected):
                create_top(text)
                if expected == "process owner":
                    expect_owner_child = True
                continue

        # ------------------------------------------------------------------
        #  Process‑Owner child (single line)
        # ------------------------------------------------------------------
        if expect_owner_child:
            child = {"heading": text, "content": [], "children": []}
            current_top["children"].append(child)
            current_sub, current_subsub = child, None
            expect_owner_child = False
            continue

        # ------------------------------------------------------------------
        #  Process Designee start (“Process Designee:” bold line)
        # ------------------------------------------------------------------
        if PROCESS_DESIGNEE_PATTERN.match(low):
            create_top("Process Designees")
            if ":" in stripped:
                after = stripped.split(":", 1)[1].strip()
                if after:
                    designee_buffer.append(after)
            capturing_designee_block = True
            continue

        # ------------------------------------------------------------------
        #  DEFINITIONS: every bold line w/ colon is new child
        # ------------------------------------------------------------------
        if (
            current_top
            and normalize(strip_numeric_prefix(current_top["heading"])).startswith("definitions")
        ):
            is_new_top = (
                current_sub is None
                and SINGLE_LEVEL_PATTERN.match(raw)
                and not SUBCLAUSE_PATTERN.match(raw)
            ) or (
                next_top_idx < len(TOP_LEVEL_SEQUENCE)
                and low.startswith(TOP_LEVEL_SEQUENCE[next_top_idx])
            )
            if not is_new_top:
                if ":" in text:
                    child = {"heading": text, "content": [], "children": []}
                    current_top["children"].append(child)
                    current_sub, current_subsub = child, None
                else:
                    if current_sub:
                        current_sub["heading"] += " " + text
                    else:
                        child = {"heading": text, "content": [], "children": []}
                        current_top["children"].append(child)
                        current_sub = child
                continue

        # ------------------------------------------------------------------
        #  POLICY REFERENCE: every bold line ⇒ new child; plain lines attach
        # ------------------------------------------------------------------
        if (
            current_top
            and normalize(strip_numeric_prefix(current_top["heading"])).startswith("policy reference")
        ):
            is_new_top = (
                current_sub is None
                and SINGLE_LEVEL_PATTERN.match(raw)
                and not SUBCLAUSE_PATTERN.match(raw)
            ) or (
                next_top_idx < len(TOP_LEVEL_SEQUENCE)
                and low.startswith(TOP_LEVEL_SEQUENCE[next_top_idx])
            )
            if not is_new_top:
                if is_bold:
                    child = {"heading": text, "content": [], "children": []}
                    current_top["children"].append(child)
                    current_sub = child
                else:
                    if current_sub:
                        current_sub["content"].append(text)
                    else:
                        child = {"heading": text, "content": [text], "children": []}
                        current_top["children"].append(child)
                        current_sub = child
                continue

        # ------------------------------------------------------------------
        #  Sub‑sub‑clause x.x.x (bold)
        # ------------------------------------------------------------------
        if is_bold and SUBSUBCLAUSE_PATTERN.match(raw) and current_sub:
            child = {"heading": text, "content": [], "children": []}
            current_sub["children"].append(child)
            current_subsub = child
            continue

        # ------------------------------------------------------------------
        #  Sub‑clause x.x (bold) – treat as content under Records
        # ------------------------------------------------------------------
        if is_bold and SUBCLAUSE_PATTERN.match(raw):
            if current_top and normalize(strip_numeric_prefix(current_top["heading"])) == "records":
                current_top["content"].append(text)
                continue
            if current_top:
                child = {"heading": text, "content": [], "children": []}
                current_top["children"].append(child)
                current_sub, current_subsub = child, None
            continue

        # ------------------------------------------------------------------
        #  Other bold line
        # ------------------------------------------------------------------
        if is_bold:
            if current_top and normalize(strip_numeric_prefix(current_top["heading"])) == "records":
                current_top["content"].append(text)
                continue
            if current_subsub:
                current_subsub["content"].append(text)
            elif current_sub:
                current_sub["content"].append(text)
            else:
                child = {"heading": text, "content": [], "children": []}
                current_top["children"].append(child)
                current_sub = child
            continue

        # ------------------------------------------------------------------
        #  Plain text – attach to deepest active node
        # ------------------------------------------------------------------
        if current_subsub:
            current_subsub["content"].append(text)
        elif current_sub:
            current_sub["content"].append(text)
        elif current_top:
            current_top["content"].append(text)

    # ------------------------------------------------------------------
    #  Flush any buffered designees
    # ------------------------------------------------------------------
    if capturing_designee_block and current_top and designee_buffer:
        for line in designee_buffer:
            token = line.strip()
            if token:
                current_top["children"].append(
                    {"heading": token, "content": [], "children": []}
                )
        # ------------------------------------------------------------------
    

    # ------------------------------------------------------------------
    #  Revision history
    # ------------------------------------------------------------------
    rev_hist = extract_revision_history(doc)
    if rev_hist["type"] == "table":
        if current_subsub:
            current_subsub["content"] = []
        elif current_sub:
            current_sub["content"] = []
        elif current_top:
            current_top["content"] = []
        return {
            "document_title": document_title,
            "purpose_scope_block": purpose_scope_block,
            "sections": sections,
            "revision_history": rev_hist,
        }

    # Written rev history → attach to last node (cleared afterwards)
    written = {"type": "written", "content": []}
    if current_subsub:
        written["content"] = current_subsub["content"]; current_subsub["content"] = []
    elif current_sub:
        written["content"] = current_sub["content"]; current_sub["content"] = []
    elif current_top:
        written["content"] = current_top["content"]; current_top["content"] = []

    return {
        "document_title": document_title,
        "purpose_scope_block": purpose_scope_block,
        "sections": sections,
        "revision_history": written,
    }
