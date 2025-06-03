from pathlib import Path
from typing import Dict, Any
import re
import logging

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.shared import Inches  # ← NEW: for indenting paragraphs

# Import the parser
from modernizer.parser import parse_legacy_docx_by_sequence

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


def _strip_all_numbers(s: str) -> str:
    """Remove ANY leading numeric prefix like '2.', '3.1 ', or '4.1.2 '."""
    return re.sub(r"^\s*\d+(?:\.\d+)*\s*", "", s).strip()


def _find_paragraph(doc: Document, exact: str) -> Paragraph:
    """
    Return the first paragraph in doc whose stripped text exactly equals `exact`.
    If none is found, return None (with a warning).
    """
    for p in doc.paragraphs:
        if p.text.strip() == exact:
            return p
    logger.warning(f"Placeholder paragraph '{exact}' not found in template.")
    return None


def _remove_paragraph(paragraph: Paragraph) -> None:
    """
    Remove a paragraph from the document.
    """
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = None


def _insert_paragraph_after(
    paragraph: Paragraph,
    text: str = "",
    style: str = None,
    indent_level: int = 0
) -> Paragraph:
    """
    Create a new paragraph immediately after `paragraph`, setting its text
    (and optionally its style). Returns the new Paragraph.

    indent_level:
      • 0 → no indent (default)
      • 1 → 0.3" left indent
      • 2 → 0.6" left indent
    """
    text = re.sub(r"\s+", " ", text).strip()
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)

    if style:
        new_para.style = style
    if text:
        new_para.text = text

    if indent_level:
        new_para.paragraph_format.left_indent = Inches(0.3 * indent_level)

    return new_para


def _insert_definition_after(
    paragraph: Paragraph,
    full_text: str
) -> Paragraph:
    """
    Insert a new paragraph after `paragraph`, bolding everything up to the first colon,
    then leaving the remainder normal. Returns the new Paragraph.
    """
    full_text = re.sub(r"\s+", " ", full_text).strip()
    if ":" in full_text:
        term, rest = full_text.split(":", 1)
        term = term.strip() + ":"
        rest = rest.strip()
    else:
        term = full_text
        rest = ""

    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)

    run1 = new_para.add_run(term)
    run1.bold = True
    if rest:
        new_para.add_run(" " + rest)

    return new_para


def _clear_table(table: Table) -> None:
    """
    Remove all rows from a table except the first (header) row.
    """
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[-1]._tr)


def normalize(text: str) -> str:
    return text.strip().lower()


def strip_numeric_prefix(text: str) -> str:
    m = re.match(r"^[\d\.]+\s*(.*)$", text.strip())
    return (m.group(1).strip() if m else text.strip())


def populate_template(
    parsed: Dict[str, Any],
    template_path: Path,
    output_path: Path
) -> None:
    """
    Open the template, then fill in each section based on the parsed structure.
    1) PURPOSE + SCOPE combined → fill under "Purpose and Scope"
    2) Remove the placeholders "PURPOSE:" and "SCOPE:" entirely
    3) REVISION HISTORY
    4) PROCESS OWNER / PROCESS DESIGNEES
    5) DEFINITIONS
    6) PROCEDURES  
    7) REFERENCES (including "Related Documents" sub‐headings)
    8) RECORDS     
    9) POLICY REFERENCE (skip/remove if not present)
    """
    doc = Document(str(template_path))

    # ------------------------------------------------------------------
    # 1) PURPOSE + SCOPE combined → under "Purpose and Scope"
    # ------------------------------------------------------------------
    p_header = _find_paragraph(doc, "Purpose and Scope")
    if p_header and parsed.get("purpose_scope_block"):
        last_p = p_header
        for line in parsed["purpose_scope_block"].splitlines():
            if line.strip():
                last_p = _insert_paragraph_after(last_p, text=line)
    else:
        if not p_header:
            logger.warning("No 'Purpose and Scope' placeholder found; skipping insertion.")
        if not parsed.get("purpose_scope_block"):
            logger.info("No Purpose/Scope block parsed; leaving template blank under that heading.")

    # After inserting, remove any standalone placeholders
    to_remove = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt in {"PURPOSE:", "Purpose:", "SCOPE:", "Scope"}:
            to_remove.append(p)
    for p in to_remove:
        _remove_paragraph(p)

    # ------------------------------------------------------------------
    # 2) REVISION HISTORY
    # ------------------------------------------------------------------
    #---p_rev_heading = _find_paragraph(doc, "Revision History")
    #existing_table = None
    #if p_rev_heading:
        #rev_p = p_rev_heading._p
        #nxt = rev_p.getnext()
        #if nxt is not None and nxt.tag == qn("w:tbl"):
            #existing_table = Table(nxt, doc)

        #rev_data = parsed.get("revision_history", {})
        #if rev_data.get("type") == "table" and existing_table:
            #rows = rev_data["rows"]
            #_clear_table(existing_table)
            #for row in rows:
                #new_row = existing_table.add_row()
                #for idx, cell_text in enumerate(row):
                    #new_row.cells[idx].text = re.sub(r"\s+", " ", cell_text).strip()
        #elif rev_data.get("type") == "written" and p_rev_heading:
            #if existing_table:
                #existing_table._tbl.getparent().remove(existing_table._tbl)
            #last_para = p_rev_heading
            #for line in rev_data.get("content", []):
                #if line.strip():
                    #last_para = _insert_paragraph_after(last_para, text=line)
        #else:
            #if existing_table:
                #existing_table._tbl.getparent().remove(existing_table._tbl)
    #else:
        #logger.warning("No 'Revision History' placeholder found; skipping revision insertion.")

    # ------------------------------------------------------------------
    # 3) PROCESS OWNER / PROCESS DESIGNEES  (2×2 table)
    # ------------------------------------------------------------------
    owner_name = ""
    designee_names = ""
    for sec in parsed["sections"]:
        head = normalize(strip_numeric_prefix(sec["heading"]))
        if head.startswith("process owner") and sec.get("children"):
            owner_name = sec["children"][0]["heading"].strip()
        elif head.startswith("process designees") and sec.get("children"):
            designee_names = ", ".join([c["heading"].strip() for c in sec["children"]])

    for tbl in doc.tables:
        hdr = tbl.cell(0, 0).text.strip().lower()
        if hdr.startswith("process owner"):
            tbl.cell(0, 1).text = owner_name
        elif hdr.startswith("process designees"):
            tbl.cell(0, 1).text = designee_names

    # ------------------------------------------------------------------
    # 4) DEFINITIONS
    # ------------------------------------------------------------------
    p_def_heading = _find_paragraph(doc, "Definitions")
    defs_node = next(
        (sec for sec in parsed["sections"]
         if "definitions" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )
    if p_def_heading and defs_node:
        last_para = p_def_heading
        for child in defs_node.get("children", []):
            full_text = child["heading"]
            if full_text.strip():
                last_para = _insert_definition_after(last_para, full_text=full_text)
                for cont in child.get("content", []):
                    for subline in cont.splitlines():
                        if subline.strip():
                            last_para = _insert_paragraph_after(last_para, text=subline)
    else:
        if not p_def_heading:
            logger.warning("No 'Definitions' placeholder found; skipping definitions.")
        if defs_node is None:
            logger.info("No 'Definitions' section in parsed doc; leaving placeholder blank.")

    # ------------------------------------------------------------------
    # 5) PROCEDURES  (4.x → Heading 2 auto‐numbered as “2.x”)
    # ------------------------------------------------------------------
    p_proc_heading = _find_paragraph(doc, "Procedures")
    proc_node = next(
        (sec for sec in parsed["sections"]
         if "procedures" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )
    if p_proc_heading and proc_node:
        last_para = p_proc_heading
        for child in proc_node.get("children", []):
            raw = re.sub(r"\s+", " ", child["heading"]).strip()
            just_text = re.sub(r"^\s*4\.\d+\s*", "", raw)
            last_para = _insert_paragraph_after(
                last_para,
                text=just_text,
                style="Heading 2"
            )
            # Insert each body line under Heading 2 with indent_level=1
            for line in child.get("content", []):
                for subline in line.splitlines():
                    if subline.strip():
                        last_para = _insert_paragraph_after(
                            last_para,
                            text=subline,
                            indent_level=1  # ← NEW
                        )
            for subsub in child.get("children", []):
                raw_ss = re.sub(r"\s+", " ", subsub["heading"]).strip()
                just_ss = re.sub(r"^\s*4\.\d+\.(\d+)\s*", r"\1 ", raw_ss)
                last_para = _insert_paragraph_after(
                    last_para,
                    text=just_ss,
                    style="Heading 3"
                )
                # Insert each body line under Heading 3 with indent_level=2
                for line2 in subsub.get("content", []):
                    for subline2 in line2.splitlines():
                        if subline2.strip():
                            last_para = _insert_paragraph_after(
                                last_para,
                                text=subline2,
                                indent_level=2  # ← NEW
                            )
    else:
        if not p_proc_heading:
            logger.warning("No 'Procedures' placeholder found; skipping procedures.")
        if proc_node is None:
            logger.info("No 'Procedures' section in parsed doc; leaving placeholder blank.")

    # ------------------------------------------------------------------
    # 6) REFERENCES  (5.x → Heading 2 auto‐numbered as “3.x”)
    #    Also insert any "Related Documents" children as sub‐headings here.
    # ------------------------------------------------------------------
    p_ref_heading = _find_paragraph(doc, "References")
    ref_node = next(
        (sec for sec in parsed["sections"]
         if "references" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )
    related_node = next(
        (sec for sec in parsed["sections"]
         if "related documents" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )

    if p_ref_heading:
        last_para = p_ref_heading

        # Insert “References” children
        if ref_node:
            for child in ref_node.get("children", []):
                raw = re.sub(r"\s+", " ", child["heading"]).strip()
                just_text = re.sub(r"^\s*5\.\d+\s*", "", raw)
                last_para = _insert_paragraph_after(
                    last_para,
                    text=just_text,
                    style="Heading 2"
                )
                # Insert each body line under Heading 2 with indent_level=1
                for line in child.get("content", []):
                    for subline in line.splitlines():
                        if subline.strip():
                            last_para = _insert_paragraph_after(
                                last_para,
                                text=subline,
                                indent_level=1  # ← NEW
                            )
            # Insert any loose lines under "References" with indent_level=1
            for line in ref_node.get("content", []):
                for subline in line.splitlines():
                    if subline.strip():
                        last_para = _insert_paragraph_after(
                            last_para,
                            text=subline,
                            indent_level=1  # ← NEW
                        )
        else:
            logger.info("No 'References' section in parsed doc; leaving placeholder blank.")

        # Insert “Related Documents” if present
        if related_node:
            for child in related_node.get("children", []):
                raw = re.sub(r"\s+", " ", child["heading"]).strip()
                last_para = _insert_paragraph_after(
                    last_para,
                    text=raw,
                    style="Heading 2"
                )
                for line in child.get("content", []):
                    for subline in line.splitlines():
                        if subline.strip():
                            last_para = _insert_paragraph_after(
                                last_para,
                                text=subline,
                                indent_level=1  # ← NEW
                            )
    else:
        logger.warning("No 'References' placeholder found; skipping references/related documents.")

    # ------------------------------------------------------------------
    # 7) RECORDS  (6.x → Heading 2 auto‐numbered as “4.x”)
    # ------------------------------------------------------------------
    p_rec = _find_paragraph(doc, "Records")
    records_node = next(
        (s for s in parsed["sections"]
         if normalize(strip_numeric_prefix(s["heading"])).startswith("records")),
        None
    )
    if p_rec and records_node:
        last = p_rec
        # Child headings first
        for child in records_node["children"]:
            txt = _strip_all_numbers(child["heading"])
            last = _insert_paragraph_after(
                last,
                text=txt,
                style="Heading 2"
            )
            # Indent each child line under Heading 2 by one level
            for ln in child.get("content", []):
                if ln.strip():
                    last = _insert_paragraph_after(
                        last,
                        text=ln,
                        indent_level=1  # ← NEW
                    )
        # Then any loose content lines under "Records" with indent_level=1
        for ln in records_node.get("content", []):
            if ln.strip():
                last = _insert_paragraph_after(
                    last,
                    text=ln,
                    indent_level=1  # ← NEW
                )
    else:
        if not p_rec:
            logger.warning("No 'Records' placeholder found; skipping records.")
        if records_node is None:
            logger.info("No 'Records' section in parsed doc; leaving placeholder blank.")

                  # ------------------------------------------------------------------
    # 8) POLICY REFERENCE  (7.x → Heading 2 auto‑numbered as “5.x”)
    #    We flatten every descendant into plain indented lines.
    # ------------------------------------------------------------------
    pol_node = next(
        (sec for sec in parsed["sections"]
         if "policy reference" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )
    p_pol_heading = _find_paragraph(doc, "Policy Reference")

    if p_pol_heading and pol_node:
        last_para = p_pol_heading

        flat_lines = []

        # -- helper to walk one level of descendants -------------------
        def collect(node):
            """Append node heading + its content lines to flat_lines."""
            head_txt = _strip_all_numbers(node["heading"]).strip()
            if head_txt:
                flat_lines.append(head_txt)

            # direct content
            for ln in node.get("content", []):
                for subln in ln.splitlines():
                    if subln.strip():
                        flat_lines.append(subln.strip())

            # one more level (grand‑children) — rare but safe
            for gchild in node.get("children", []):
                g_head = _strip_all_numbers(gchild["heading"]).strip()
                if g_head:
                    flat_lines.append(g_head)
                for ln2 in gchild.get("content", []):
                    for subln2 in ln2.splitlines():
                        if subln2.strip():
                            flat_lines.append(subln2.strip())

        # collect every child (and its descendants)
        for child in pol_node.get("children", []):
            collect(child)

        # loose lines on the Policy‑Reference node itself
        for ln in pol_node.get("content", []):
            for subln in ln.splitlines():
                if subln.strip():
                    flat_lines.append(subln.strip())

        # insert them in order, indented 0.25"
        for txt in flat_lines:
            last_para = _insert_paragraph_after(
                last_para,
                text=txt,
                indent_level=1
            )

    else:
        # Placeholder missing OR legacy doc had no Policy Reference
        if p_pol_heading and not pol_node:
            logger.info("No 'Policy Reference' in parsed doc; removing placeholder.")
            _remove_paragraph(p_pol_heading)
        elif pol_node and not p_pol_heading:
            logger.warning(
                "Parsed a 'Policy Reference' section but template has no placeholder. "
                "Skipping Policy Reference insertion."
            )




    # ------------------------------------------------------------------
    # 9) SAVE (with "-copy" logic if the file already exists)
    # ------------------------------------------------------------------
    if output_path.exists():
        base = output_path.stem
        ext = output_path.suffix
        parent = output_path.parent
        counter = 1
        candidate = parent / f"{base}-copy{ext}"
        while candidate.exists():
            counter += 1
            candidate = parent / f"{base}-copy{counter}{ext}"
        output_path = candidate
        logger.info(f"Output file exists; saving as '{output_path.name}' instead.")

    doc.save(str(output_path))
    logger.info(f"Saved modernized document to: {output_path}")
