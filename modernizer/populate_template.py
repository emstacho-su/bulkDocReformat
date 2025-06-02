# modernizer/populate_template.py

from pathlib import Path
from typing import Dict, Any, List
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

# Import your parser
from modernizer.parser import parse_legacy_docx_by_sequence


def _find_paragraph(doc: Document, startswith: str) -> Paragraph:
    """
    Return the first paragraph in doc whose text begins with 'startswith'. 
    If none found, raises ValueError.
    """
    for p in doc.paragraphs:
        if p.text.strip().startswith(startswith):
            return p
    raise ValueError(f"Cannot find paragraph starting with '{startswith}'.")


def _insert_after_paragraph(paragraph: Paragraph, new_element) -> Paragraph:
    """
    Insert an XML element (paragraph or table) immediately after `paragraph` 
    in the document, and return the newly inserted object as a Paragraph or Table.
    """
    p = paragraph._p  # the <w:p> element
    new = new_element._element if isinstance(new_element, Paragraph) else new_element._tbl
    p.addnext(new)
    return new_element


def _add_table_after(paragraph: Paragraph, rows: List[List[str]]) -> Table:
    """
    Immediately after `paragraph`, insert a new table with the given 2D list of strings.
    Returns the new Table object.
    """
    tbl = paragraph.insert_paragraph_after().add_table(rows=rows, cols=len(rows[0]))
    return tbl


def _clear_table(table: Table) -> None:
    """
    Remove all rows from an existing table, leaving just one blank row.
    """
    # delete all rows except the first
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[-1]._tr)


def populate_template(
    parsed: Dict[str, Any],
    template_path: Path,
    output_path: Path
) -> None:
    """
    Open the template at `template_path`, then:

      1. Under the existing “Purpose:” heading, insert all lines from 
         parsed["sections"] for “purpose” as normal paragraphs.

      2. Under “Scope:”, insert all lines from parsed["sections"] for “scope”.

      3. Under “Revision History”, replace the existing placeholder table 
         (if any) with a properly formatted table built from parsed["revision_history"].

      4. Fill the 2×2 “Process Owner / Process Designees” table:

         • Locate the table whose first cell literally contains “Process Owner/Authorized By:”  
         • In cell (0,1), put the single child of parsed[“process owner”].  
         • Locate the table whose first cell literally contains “Process Designees:”  
         • In cell (0,1), put the paragraph(s) that follow “process designee” in the parsed data.

      5. Leave placeholders for the TOC page breaks; we rely on Word’s built‐in TOC 
         feature rather than generating field codes programmatically. At least ensure 
         that the “Table of Contents” heading remains in place.

      6. Under the template’s “1    Definitions” heading, insert each definition child 
         (from parsed["definitions"]) as a separate paragraph (numbering and bolding is 
         already handled by Word’s style). 

      7. Under “2    Procedures,” insert subclauses and sub‐subclauses from parsed["procedures"], 
         adjusting the leading number from “4.x” to “2.x” to match the new template’s numbering. 
         Every x.x (formerly 4.x) becomes “2.x”, every x.x.x (formerly 4.x.x) becomes “2.x.x”, etc.

      8. Under “3    References,” insert parsed["references"] (adjust old “5.x” → “3.x”).

      9. Under “4    Records,” insert parsed["records"] (old “6.x” → “4.x”).

     10. Under “5    Policy Reference,” insert parsed["policy reference"] (old “7.x” → “5.x”).

    Finally, save to `output_path`.
    """
    doc = Document(str(template_path))

    # 1) Populate Purpose
    try:
        p_purpose = _find_paragraph(doc, "Purpose:")
    except ValueError:
        raise

    # Find parsed “purpose” node:
    purpose_node = next((sec for sec in parsed["sections"] 
                         if normalize(strip_numeric_prefix(sec["heading"])).startswith("purpose")), None)
    if purpose_node:
        for line in purpose_node.get("content", []):
            if line.strip():
                p_purpose.insert_paragraph_after(line)

    # 2) Populate Scope
    try:
        p_scope = _find_paragraph(doc, "Scope:")
    except ValueError:
        raise

    scope_node = next((sec for sec in parsed["sections"] 
                       if normalize(strip_numeric_prefix(sec["heading"])).startswith("scope")), None)
    if scope_node:
        for line in scope_node.get("content", []):
            if line.strip():
                p_scope.insert_paragraph_after(line)

    # 3) Populate Revision History Table
    try:
        p_rev_heading = _find_paragraph(doc, "Revision History")
    except ValueError:
        raise

    # If the template already has a table right after “Revision History,” remove its rows
    # and replace with the parsed rows. Otherwise, just insert a new table.
    # We look for a table immediately following p_rev_heading.
    insert_point = None
    for idx, elem in enumerate(doc.paragraphs):
        if elem is p_rev_heading:
            insert_point = idx + 1
            break

    # Check if there’s a table in doc._body._body—next sibling
    existing_table = None
    if insert_point is not None and insert_point < len(doc.paragraphs):
        # Use the underlying XML siblings to find a TABLE right after p_rev_heading._p
        rev_p = p_rev_heading._p
        nxt = rev_p.getnext()
        if nxt is not None and nxt.tag == qn("w:tbl"):
            existing_table = Table(nxt, doc)

    rev_data = parsed.get("revision_history", {})
    if rev_data.get("type") == "table":
        rows = rev_data["rows"]
        if existing_table:
            # Clear it, then fill new rows
            _clear_table(existing_table)
            for row in rows:
                existing_table.add_row().cells[:] = [cell_text for cell_text in row]
        else:
            # Insert a brand‐new table after the heading
            table = p_rev_heading.insert_paragraph_after().add_table(rows=rows, cols=len(rows[0]))
            # Already filled by add_table
    else:
        # Written revision history: place each line as a new paragraph under the heading
        if existing_table:
            # Remove the table entirely
            existing_table._tbl.getparent().remove(existing_table._tbl)
        # Then simply write each line from rev_data["content"]
        for line in rev_data.get("content", []):
            if line.strip():
                p_rev_heading.insert_paragraph_after(line)

    # 4) Populate Process Owner/Designees 2×2 table(s)
    proc_owner_child = None
    proc_designee_child = None

    # Find parsed nodes
    po_node = next((sec for sec in parsed["sections"] 
                    if normalize(strip_numeric_prefix(sec["heading"])).startswith("process owner")), None)
    if po_node and po_node.get("children"):
        # The first child under “Process Owner”
        proc_owner_child = po_node["children"][0]["heading"]

    pd_node = next((sec for sec in parsed["sections"] 
                    if normalize(strip_numeric_prefix(sec["heading"])).startswith("process designee")), None)
    if pd_node and pd_node.get("children"):
        # Designee block was lumped as single child under “Process Designee”
        proc_designee_child = pd_node["children"][0]["heading"]

    # Now search all doc.tables for the 2×2
    for table in doc.tables:
        first_cell_text = table.cell(0, 0).text.strip()
        if first_cell_text.startswith("Process Owner/Authorized By"):
            if proc_owner_child:
                # Place in cell(0,1)
                table.cell(0, 1).text = proc_owner_child
        elif first_cell_text.startswith("Process Designees"):
            if proc_designee_child:
                table.cell(0, 1).text = proc_designee_child

    # 5) Table of Contents: 
    # We assume the template already has a “Table of Contents” heading and a placeholder below it.
    # python-docx cannot insert the TOC field automatically, so we leave the placeholder in place.
    # After the user opens this file in Word, they must right-click on the TOC and “Update Field.”
    # We won’t modify anything here.

    # 6) Populate “1    Definitions”
    try:
        p_def_heading = _find_paragraph(doc, "1    Definitions")
    except ValueError:
        raise

    defs_node = next((sec for sec in parsed["sections"] 
                      if normalize(strip_numeric_prefix(sec["heading"])).startswith("definitions")), None)
    if defs_node:
        for child in defs_node.get("children", []):
            # Each child["heading"] is a separate definition
            if child["heading"].strip():
                p_def_heading.insert_paragraph_after(child["heading"])

    # 7) Populate “2    Procedures” with remapped numbering
    try:
        p_proc_heading = _find_paragraph(doc, "2    Procedures")
    except ValueError:
        raise

    proc_node = next((sec for sec in parsed["sections"] 
                      if normalize(strip_numeric_prefix(sec["heading"])).startswith("procedures")), None)
    if proc_node:
        for child in proc_node.get("children", []):
            raw = child["heading"].strip()
            # Remap “4.x” → “2.x”; “4.x.x” → “2.x.x”
            new_heading = re.sub(r"^\s*4\.", "2.", raw)
            para = p_proc_heading.insert_paragraph_after(new_heading)
            # Now write content of that subclause
            for line in child.get("content", []):
                if line.strip():
                    para.insert_paragraph_after(line)
            # Recurse sub‐subclause under this child, if any
            for subsub in child.get("children", []):
                raw_ss = subsub["heading"].strip()
                new_ss = re.sub(r"^\s*4\.([0-9]+)\.", r"2.\1.", raw_ss)
                ss_para = para.insert_paragraph_after(new_ss)
                for line2 in subsub.get("content", []):
                    if line2.strip():
                        ss_para.insert_paragraph_after(line2)

    # 8) Populate “3    References” (remap “5.x” → “3.x”)
    try:
        p_ref_heading = _find_paragraph(doc, "3    References")
    except ValueError:
        raise

    ref_node = next((sec for sec in parsed["sections"] 
                     if normalize(strip_numeric_prefix(sec["heading"])).startswith("references")), None)
    if ref_node:
        for child in ref_node.get("children", []):
            raw = child["heading"].strip()
            new_heading = re.sub(r"^\s*5\.", "3.", raw)
            para = p_ref_heading.insert_paragraph_after(new_heading)
            for line in child.get("content", []):
                if line.strip():
                    para.insert_paragraph_after(line)

    # 9) Populate “4    Records” (remap “6.x” → “4.x”)
    try:
        p_rec_heading = _find_paragraph(doc, "4    Records")
    except ValueError:
        raise

    rec_node = next((sec for sec in parsed["sections"] 
                     if normalize(strip_numeric_prefix(sec["heading"])).startswith("records")), None)
    if rec_node:
        for child in rec_node.get("children", []):
            raw = child["heading"].strip()
            new_heading = re.sub(r"^\s*6\.", "4.", raw)
            para = p_rec_heading.insert_paragraph_after(new_heading)
            for line in child.get("content", []):
                if line.strip():
                    para.insert_paragraph_after(line)

    # 10) Populate “5    Policy Reference” (remap “7.x” → “5.x”)
    try:
        p_pol_heading = _find_paragraph(doc, "5    Policy Reference")
    except ValueError:
        raise

    pol_node = next((sec for sec in parsed["sections"] 
                     if normalize(strip_numeric_prefix(sec["heading"])).startswith("policy reference")), None)
    if pol_node:
        for child in pol_node.get("children", []):
            raw = child["heading"].strip()
            new_heading = re.sub(r"^\s*7\.", "5.", raw)
            para = p_pol_heading.insert_paragraph_after(new_heading)
            for line in child.get("content", []):
                if line.strip():
                    para.insert_paragraph_after(line)

    # 11) Save new document
    doc.save(str(output_path))


# Helper utilities used above:
def normalize(text: str) -> str:
    return text.strip().lower()


def strip_numeric_prefix(text: str) -> str:
    m = re.match(r"^[\d\.]+\s*(.*)$", text.strip())
    return (m.group(1).strip() if m else text.strip())
