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
    "purpose",  # only used for ordering, but we now handle it specially
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

# Regex to match subclause pattern "x.x" (e.g. "4.1", "5.2", "6.3")
SUBCLAUSE_PATTERN = re.compile(r"^\s*\d+\.\d+")

# Regex to match sub‐subclause pattern "x.x.x" (e.g. "4.1.1", "5.2.3")
SUBSUBCLAUSE_PATTERN = re.compile(r"^\s*\d+\.\d+\.\d+")

# Regex to match single‐level prefix "7. Something" but not "7.1"
SINGLE_LEVEL_PATTERN = re.compile(r"^(\d+)\.\s+")

# Case‐insensitive match for “purpose and scope”
PURPOSE_AND_SCOPE_PATTERN = re.compile(r"^purpose\s+and\s+scope", re.IGNORECASE)

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
    Parses a legacy .docx with these adjustments:
    
    1) Everything from immediately after “1. Purpose and Scope” up to “2. Definitions”
       is gathered into a single block.  We then split that block on the first
       occurrence of “Scope” (case‐insensitive) so that:
         - purpose_text = lines between “PURPOSE” and “SCOPE”
         - scope_text   = lines between “SCOPE” and “2. Definitions”
       These two strings are returned separately as `purpose_text` and `scope_text`.
    
    2) From “2. Definitions” onward, parsing proceeds as before––extracting top‐levels
       (Definitions, Process Owner, etc.), subclauses, sub‐subclauses, and finally
       revision history.

    3) Under “Records,” we force bold lines to attach as content, never as subclauses.

    Returns a dict containing:
      {
        "document_title": str,
        "purpose_text": str,
        "scope_text": str,
        "sections": [...],         # parsed from “2. Definitions” onward
        "revision_history": {...}
      }
    """

    # (A) Load document and collect all non‐empty paragraphs
    try:
        doc = Document(path)
    except Exception as e:
        logger.error(f"Failed to open {path.name}: {e}")
        raise

    paras = [p for p in doc.paragraphs if p.text.strip()]
    if not paras:
        logger.warning(f"No paragraphs found in {path.name}")
        return {
            "document_title": "",
            "purpose_text": "",
            "scope_text": "",
            "sections": [],
            "revision_history": {"type": "written", "content": []}
        }

    # (B) Determine Document Title = first bold paragraph, or fallback to first
    document_title = ""
    for p in paras:
        if any(run.bold for run in p.runs):
            document_title = p.text.strip()
            break
    if not document_title:
        document_title = paras[0].text.strip()
        logger.warning(f"No bold title in {path.name}; using first line")

    logger.info(f"[{path.name}] Document Title: {document_title!r}")

    # (C) Locate the indices of “1. Purpose and Scope” and “2. Definitions”
    idx_purpose_scope = None
    idx_definitions = None
    for idx, p in enumerate(paras):
        raw = p.text.strip()
        stripped = strip_numeric_prefix(raw)
        low = normalize(stripped)

        # Find “1. Purpose and Scope”
        if idx_purpose_scope is None:
            if SINGLE_LEVEL_PATTERN.match(raw):
                # If strip_numeric_prefix(raw) starts with "Purpose and Scope"
                if PURPOSE_AND_SCOPE_PATTERN.match(stripped):
                    idx_purpose_scope = idx
                    continue

        # Once we have idx_purpose_scope, find “2. Definitions”
        if idx_purpose_scope is not None and idx_definitions is None:
            if SINGLE_LEVEL_PATTERN.match(raw):
                # If strip_numeric_prefix(raw) starts with "Definitions"
                if low.startswith("definitions"):
                    idx_definitions = idx
                    break

    # If we never found “1. Purpose and Scope,” fall back to normal parsing
    if idx_purpose_scope is None or idx_definitions is None:
        logger.warning("Could not locate both '1. Purpose and Scope' and '2. Definitions'. "
                       "Skipping special-purpose handling.")
        idx_purpose_scope = None
        idx_definitions = None

    # (D) Extract purpose_block and scope_block if we have both indices
    purpose_text = ""
    scope_text = ""
    start_idx = (idx_purpose_scope + 1) if idx_purpose_scope is not None else None
    if start_idx is not None:
        block_paras = [p.text.strip() for p in paras[start_idx:idx_definitions]]
        # Find the split index where “Scope” appears
        split_idx = None
        for i, line in enumerate(block_paras):
            # normalized word starts with “scope”
            if normalize(strip_numeric_prefix(line)).startswith("scope"):
                split_idx = i
                break

        # If we found “SCOPE” somewhere
        if split_idx is not None:
            # Everything from after the bold “PURPOSE” line up to before “SCOPE”
            purpose_lines = []
            for line in block_paras[:split_idx]:
                # skip any leading “PURPOSE” label
                if normalize(strip_numeric_prefix(line)).startswith("purpose"):
                    continue
                purpose_lines.append(line)
            purpose_text = "\n".join(purpose_lines).strip()

            # Everything after “SCOPE” label
            scope_lines = []
            for line in block_paras[split_idx + 1 :]:
                scope_lines.append(line)
            scope_text = "\n".join(scope_lines).strip()
        else:
            # If “SCOPE” keyword not found, everything goes into purpose_text
            purpose_text = "\n".join(block_paras).strip()
            scope_text = ""

        logger.info(f"Extracted purpose_text: {purpose_text!r}")
        logger.info(f"Extracted scope_text:   {scope_text!r}")

    # (E) Now parse from “2. Definitions” onward
    remaining_paras = paras[idx_definitions:] if idx_definitions is not None else paras[:]

    # Redefine local lists for the remainder
    sections: List[Dict[str, Any]] = []
    current_top: Dict[str, Any] = None
    current_sub: Dict[str, Any] = None
    current_subsub: Dict[str, Any] = None
    next_top_idx = 2  # We have effectively consumed “purpose” and “scope” indices

    expect_owner_child = False
    capturing_designee_block = False
    designee_buffer: List[str] = []

    def create_top(heading: str, forced_keyword: str = None) -> None:
        nonlocal sections, current_top, current_sub, current_subsub, next_top_idx, expect_owner_child, capturing_designee_block, designee_buffer

        node = {"heading": heading, "content": [], "children": []}
        sections.append(node)
        current_top = node
        current_sub = None
        current_subsub = None
        expect_owner_child = False
        capturing_designee_block = False
        designee_buffer = []

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

    # (F) Iterate through “remaining_paras” exactly as before, with two exceptions:
    #     1) “Records” never spawns subclauses
    #     2) Multi‐line “Process Designee”
    #
    for p in remaining_paras:
        text = p.text.strip()
        if not text:
            continue

        stripped = strip_numeric_prefix(text)
        low = normalize(stripped)
        raw = text.strip()
        is_bold = any(run.bold for run in p.runs)

        # 1) If capturing a Designee block, gather until next real top or subclause
        if capturing_designee_block:
            is_new_top = (
                current_sub is None
                and SINGLE_LEVEL_PATTERN.match(raw)
                and not SUBCLAUSE_PATTERN.match(raw)
            ) or (
                next_top_idx < len(TOP_LEVEL_SEQUENCE)
                and low.startswith(TOP_LEVEL_SEQUENCE[next_top_idx])
            )
            is_procedures = (
                low.startswith("procedures")
                and next_top_idx < len(TOP_LEVEL_SEQUENCE)
                and TOP_LEVEL_SEQUENCE[next_top_idx] == "procedures"
            )
            if is_new_top or is_procedures:
                # Flush buffered names
                for name_line in designee_buffer:
                    for token in re.split(r"\s*,\s*", name_line):
                        token = token.strip()
                        if token:
                            child = {"heading": token, "content": [], "children": []}
                            current_top["children"].append(child)
                designee_buffer = []
                capturing_designee_block = False
                # fall through to re‐process this paragraph
            else:
                designee_buffer.append(text)
                continue

        # 2) SINGLE‐LEVEL numeric prefix → NEW TOP (unless “1. Purpose and Scope,” but now we’re at “2. Definitions” onward)
        if current_sub is None and SINGLE_LEVEL_PATTERN.match(raw) and not SUBCLAUSE_PATTERN.match(raw):
            # At this point, “2. Definitions” should match and create a top node
            create_top(text)
            continue

        # 3) TOP‐LEVEL keyword in sequence (e.g. “Definitions,” “Process Owner,” etc.)
        if next_top_idx < len(TOP_LEVEL_SEQUENCE):
            expected = TOP_LEVEL_SEQUENCE[next_top_idx]
            if low.startswith(expected):
                if PROCESS_DESIGNEE_PATTERN.match(low):
                    pass
                else:
                    create_top(text)
                    if low.startswith("process owner"):
                        expect_owner_child = True
                    continue

        # 4) PROCESS OWNER → next non‐empty paragraph is its single child
        if expect_owner_child:
            child = {"heading": text, "content": [], "children": []}
            current_top["children"].append(child)
            current_sub = child
            current_subsub = None
            expect_owner_child = False
            continue

        # 5) PROCESS DESIGNEE: build multi‐line children
        if PROCESS_DESIGNEE_PATTERN.match(low):
            create_top("Process Designees")
            if ":" in stripped:
                after = stripped.split(":", 1)[1].strip()
                if after:
                    designee_buffer.append(after)
            capturing_designee_block = True
            continue

        # 6) “DEFINITIONS” section:
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
                    current_sub = child
                    current_subsub = None
                else:
                    if current_sub:
                        current_sub["heading"] += " " + text
                    else:
                        child = {"heading": text, "content": [], "children": []}
                        current_top["children"].append(child)
                        current_sub = child
                        current_subsub = None
                continue

        # 7) SUB‐SUBCLAUSE “x.x.x” (only if bold, and not under “Records”)
        if is_bold and SUBSUBCLAUSE_PATTERN.match(raw) and current_sub:
            child = {"heading": text, "content": [], "children": []}
            current_sub["children"].append(child)
            current_subsub = child
            continue

        # 8) SUBCLAUSE “x.x” (only if bold; under “Records,” attach as content)
        if is_bold and SUBCLAUSE_PATTERN.match(raw):
            if (
                current_top
                and normalize(strip_numeric_prefix(current_top["heading"])) == "records"
            ):
                current_top["content"].append(text)
                continue
            if current_top:
                child = {"heading": text, "content": [], "children": []}
                current_top["children"].append(child)
                current_sub = child
                current_subsub = None
            continue

        # 9) Any other bold:
        if is_bold:
            if (
                current_top
                and normalize(strip_numeric_prefix(current_top["heading"])) == "records"
            ):
                current_top["content"].append(text)
                continue
            if current_subsub:
                current_subsub["content"].append(text)
                continue
            if current_sub:
                current_sub["content"].append(text)
                continue
            if current_top:
                child = {"heading": text, "content": [], "children": []}
                current_top["children"].append(child)
                current_sub = child
                current_subsub = None
            continue

        # 10) Non‐bold → attach as content under the deepest active node
        if current_subsub:
            current_subsub["content"].append(text)
        elif current_sub:
            current_sub["content"].append(text)
        elif current_top:
            current_top["content"].append(text)
        else:
            # No parent
            pass

    # (G) Flush any remaining designee_buffer
    if capturing_designee_block and current_top and designee_buffer:
        for name_line in designee_buffer:
            for token in re.split(r"\s*,\s*", name_line):
                token = token.strip()
                if token:
                    child = {"heading": token, "content": [], "children": []}
                    current_top["children"].append(child)
        designee_buffer = []
        capturing_designee_block = False

    # (H) Extract revision history
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
            "purpose_text": purpose_text,
            "scope_text": scope_text,
            "sections": sections,
            "revision_history": rev_hist
        }

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

    return {
        "document_title": document_title,
        "purpose_text": purpose_text,
        "scope_text": scope_text,
        "sections": sections,
        "revision_history": written
    }
