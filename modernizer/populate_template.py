# modernizer/populate_template.py

from pathlib import Path
from typing import Dict, Any
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

# Import the parser
from modernizer.parser import parse_legacy_docx_by_sequence

def _strip_all_numbers(s: str) -> str:
    """Remove ANY leading numeric prefix like '2.', '3.1 ', or '4.1.2 '."""
    return re.sub(r"^\s*\d+(?:\.\d+)*\s*", "", s).strip()

def _find_paragraph(doc: Document, exact: str) -> Paragraph:
    """
    Return the first paragraph in doc whose stripped text exactly equals `exact`.
    Raises ValueError if none is found.
    """
    for p in doc.paragraphs:
        if p.text.strip() == exact:
            return p
    raise ValueError(f"Cannot find paragraph with text '{exact}'.")

def _insert_paragraph_after(
    paragraph: Paragraph,
    text: str = "",
    style: str = None
) -> Paragraph:
    """
    Create a new paragraph immediately after `paragraph`, setting its text
    (and optionally its style). Returns the new Paragraph.
    """
    text = re.sub(r"\s+", " ", text).strip()
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)

    if style:
        new_para.style = style
    if text:
        new_para.text = text

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

    1) PURPOSE  → “PURPOSE” top‐level node
    2) SCOPE    → “SCOPE” top‐level node
    3) REVISION HISTORY
    4) PROCESS OWNER / PROCESS DESIGNEES
    5) DEFINITIONS
    6) PROCEDURES  (4.x → Heading 2 auto‐numbered as “2.x”)
    7) REFERENCES  (5.x → Heading 2 auto‐numbered as “3.x”)
    8) RECORDS     (6.x → Heading 2 auto‐numbered as “4.x”)
    9) POLICY REFERENCE (7.x → Heading 2 auto‐numbered as “5.x”)

    Debug prints remain for “PURPOSE” and “SCOPE” so you see what got parsed.
    """
    doc = Document(str(template_path))

    # ------------------------------------------------------------------
    # 1) PURPOSE  (placeholder paragraph literally contains "Purpose:")
    # ------------------------------------------------------------------
    p_purpose = _find_paragraph(doc, "Purpose:")
    last_p = p_purpose
    for line in parsed["purpose_text"].splitlines():
        if line.strip():
            last_p = _insert_paragraph_after(last_p, text=line)

    # ------------------------------------------------------------------
    # 2) SCOPE  (placeholder paragraph literally contains "Scope:")
    # ------------------------------------------------------------------
    p_scope = _find_paragraph(doc, "Scope:")
    last_s = p_scope
    for line in parsed["scope_text"].splitlines():
        if line.strip():
            last_s = _insert_paragraph_after(last_s, text=line)

    # ------------------------------------------------------------------
    # 3) REVISION HISTORY
    # ------------------------------------------------------------------
    p_rev_heading = _find_paragraph(doc, "Revision History")
    existing_table = None
    rev_p = p_rev_heading._p
    nxt = rev_p.getnext()
    if nxt is not None and nxt.tag == qn("w:tbl"):
        existing_table = Table(nxt, doc)

    rev_data = parsed.get("revision_history", {})
    if rev_data.get("type") == "table":
        rows = rev_data["rows"]
        if existing_table:
            _clear_table(existing_table)
            for row in rows:
                new_row = existing_table.add_row()
                for idx, cell_text in enumerate(row):
                    new_row.cells[idx].text = re.sub(r"\s+", " ", cell_text).strip()
        else:
            tbl_para = _insert_paragraph_after(p_rev_heading)
            new_tbl = doc.add_table(rows=0, cols=len(rows[0]))
            rev_p.addnext(new_tbl._tbl)
            for row in rows:
                new_row = new_tbl.add_row()
                for idx, cell_text in enumerate(row):
                    new_row.cells[idx].text = re.sub(r"\s+", " ", cell_text).strip()
    else:
        if existing_table:
            existing_table._tbl.getparent().remove(existing_table._tbl)
        last_para = p_rev_heading
        for line in rev_data.get("content", []):
            if line.strip():
                last_para = _insert_paragraph_after(
                    last_para,
                    text=line
                )

    # ------------------------------------------------------------------
    # 4) PROCESS OWNER / PROCESS DESIGNEES  (2×2 table)
    # ------------------------------------------------------------------
    owner_name = ""
    designee_names = ""

    for sec in parsed["sections"]:
        head = normalize(strip_numeric_prefix(sec["heading"]))
        if head.startswith("process owner") and sec["children"]:
            owner_name = sec["children"][0]["heading"].strip()
        elif head.startswith("process designee") and sec["children"]:
            designee_names = sec["children"][0]["heading"].strip()

    for tbl in doc.tables:
        hdr = tbl.cell(0, 0).text.strip().lower()
        if hdr.startswith("process owner"):
            tbl.cell(0, 1).text = owner_name
        elif hdr.startswith("process designee"):
            tbl.cell(0, 1).text = designee_names

    # ------------------------------------------------------------------
    # 5) DEFINITIONS
    # ------------------------------------------------------------------
    p_def_heading = _find_paragraph(doc, "Definitions")
    defs_node = next(
        (sec for sec in parsed["sections"]
         if "definitions" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )
    if defs_node:
        last_para = p_def_heading
        for child in defs_node.get("children", []):
            full_text = child["heading"]
            if full_text.strip():
                last_para = _insert_definition_after(
                    last_para,
                    full_text=full_text
                )
                for cont in child.get("content", []):
                    for subline in cont.splitlines():
                        if subline.strip():
                            last_para = _insert_paragraph_after(
                                last_para,
                                text=subline
                            )

    # ------------------------------------------------------------------
    # 6) PROCEDURES  (4.x → Heading 2 auto‐numbered as “2.x”)
    # ------------------------------------------------------------------
    p_proc_heading = _find_paragraph(doc, "Procedures")
    proc_node = next(
        (sec for sec in parsed["sections"]
         if "procedures" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )
    if proc_node:
        last_para = p_proc_heading
        for child in proc_node.get("children", []):
            raw = re.sub(r"\s+", " ", child["heading"]).strip()
            just_text = re.sub(r"^\s*4\.\d+\s*", "", raw)
            last_para = _insert_paragraph_after(
                last_para,
                text=just_text,
                style="Heading 2"
            )
            for line in child.get("content", []):
                for subline in line.splitlines():
                    if subline.strip():
                        last_para = _insert_paragraph_after(
                            last_para,
                            text=subline
                        )
            for subsub in child.get("children", []):
                raw_ss = re.sub(r"\s+", " ", subsub["heading"]).strip()
                just_ss = re.sub(r"^\s*4\.\d+\.(\d+)\s*", r"\1 ", raw_ss)
                last_para = _insert_paragraph_after(
                    last_para,
                    text=just_ss,
                    style="Heading 3"
                )
                for line2 in subsub.get("content", []):
                    for subline2 in line2.splitlines():
                        if subline2.strip():
                            last_para = _insert_paragraph_after(
                                last_para,
                                text=subline2
                            )

    # ------------------------------------------------------------------
    # 7) REFERENCES  (5.x → Heading 2 auto‐numbered as “3.x”)
    # ------------------------------------------------------------------
    p_ref_heading = _find_paragraph(doc, "References")
    ref_node = next(
        (sec for sec in parsed["sections"]
         if "references" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )
    if ref_node:
        last_para = p_ref_heading
        for child in ref_node.get("children", []):
            raw = re.sub(r"\s+", " ", child["heading"]).strip()
            just_text = re.sub(r"^\s*5\.\d+\s*", "", raw)
            last_para = _insert_paragraph_after(
                last_para,
                text=just_text,
                style="Heading 2"
            )
            for line in child.get("content", []):
                for subline in line.splitlines():
                    if subline.strip():
                        last_para = _insert_paragraph_after(
                            last_para,
                            text=subline
                        )

    # ------------------------------------------------------------------
    # 8) RECORDS  (children *and* direct content)
    # ------------------------------------------------------------------
    p_rec = _find_paragraph(doc, "Records")
    records_node = next(
        (s for s in parsed["sections"]
         if normalize(strip_numeric_prefix(s["heading"])).startswith("records")),
        None
    )

    if records_node:
        last = p_rec
        # 1) child headings first (rare but supported)
        for child in records_node["children"]:
            txt = _strip_all_numbers(child["heading"])
            last = _insert_paragraph_after(last, text=txt, style="Heading 2")
            for ln in child["content"]:
                if ln.strip():
                    last = _insert_paragraph_after(last, text=ln)
        # 2) any loose content lines
        for ln in records_node["content"]:
            if ln.strip():
                last = _insert_paragraph_after(last, text=ln)

    # ------------------------------------------------------------------
    # 9) POLICY REFERENCE  (7.x → Heading 2 auto‐numbered as “5.x”)
    # ------------------------------------------------------------------
    p_pol_heading = _find_paragraph(doc, "Policy Reference")
    pol_node = next(
        (sec for sec in parsed["sections"]
         if "policy reference" in normalize(strip_numeric_prefix(sec["heading"]))),
        None
    )
    if pol_node:
        last_para = p_pol_heading
        for child in pol_node.get("children", []):
            raw = re.sub(r"\s+", " ", child["heading"]).strip()
            just_text = re.sub(r"^\s*7\.\d+\s*", "", raw)
            last_para = _insert_paragraph_after(
                last_para,
                text=just_text,
                style="Heading 2"
            )
            for line in child.get("content", []):
                for subline in line.splitlines():
                    if subline.strip():
                        last_para = _insert_paragraph_after(
                            last_para,
                            text=subline
                        )

    # ------------------------------------------------------------------
    # 10) SAVE
    # ------------------------------------------------------------------
    doc.save(str(output_path))
