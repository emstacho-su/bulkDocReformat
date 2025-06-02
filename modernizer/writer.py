# modernizer/writer.py

from pathlib import Path
from typing import Dict, Any
import re

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

# Import your parser method
from modernizer.parser import parse_legacy_docx_by_sequence


def write_new_doc(parsed: Dict[str, Any], output_path: Path) -> None:
    """
    Given the dictionary returned by `parse_legacy_docx_by_sequence(path)`,
    create a new .docx at `output_path` with this structure:

      • Title (parsed["document_title"]) as Heading 1
      • For each top‐level in parsed["sections"]:
          – Heading 2 for top["heading"]
          – Any top["content"] lines as normal paragraphs
          – For each child in top["children"]:
              • If child["heading"] matches x.x.x → Heading 4
              • Else if child["heading"] matches x.x   → Heading 3
              • Else (catch‐all)                          → Heading 3
              • Any child["content"] lines as normal paragraphs
              • Recurse one more level: any child["children"]
                  – Write as Heading 4, etc., with their content

    Styles are mapped as follows:
      • Heading 1 → 16 pt, bold, centered
      • Heading 2 → 14 pt, bold
      • Heading 3 → 12 pt, bold
      • Heading 4 → 11 pt, bold, italic
      • Body text → 11 pt, normal

    Args:
        parsed:  The dict from parse_legacy_docx_by_sequence(...)
        output_path:  Path where the new .docx should be saved
    """
    doc = Document()

    # 1) Document Title (Heading 1, centered)
    title_para = doc.add_paragraph(parsed["document_title"], style="Heading 1")
    title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # 2) Helper to write a block (node) at a given heading level
    def write_node(node: Dict[str, Any], level: int) -> None:
        """
        Recursively write a node (with keys "heading", "content", "children") at given
        heading level. level=2→Heading 2, 3→Heading 3, 4→Heading 4. After the heading, write
        node["content"] as normal paragraphs, then recurse into node["children"].
        """
        # Choose style name based on level
        if level == 1:
            style_name = "Heading 1"
        elif level == 2:
            style_name = "Heading 2"
        elif level == 3:
            style_name = "Heading 3"
        else:
            style_name = "Heading 4"

        # 2.a) Write the heading line
        heading_para = doc.add_paragraph(node["heading"], style=style_name)

        # 2.b) Write any content lines under the heading
        for line in node.get("content", []):
            if line.strip():
                body_para = doc.add_paragraph(line)
                body_para.style.font.size = Pt(11)

        # 2.c) Recurse into children
        for child in node.get("children", []):
            # Determine if this is a sub‐subclause (e.g. starts with x.x.x)
            raw = child["heading"].strip()
            if re.match(r"^\s*\d+\.\d+\.\d+", raw):
                # three‐level (x.x.x) → increase level by 1 (max 4)
                write_node(child, min(level + 1, 4))
            elif re.match(r"^\s*\d+\.\d+", raw):
                # two‐level (x.x) → level + 1 (max 4)
                write_node(child, min(level + 1, 4))
            else:
                # otherwise treat as same‐level child under top, so also Heading (level+1)
                write_node(child, min(level + 1, 4))

    # 3) Walk through each top‐level section (Heading 2)
    for top in parsed.get("sections", []):
        write_node(top, level=2)

    # 4) Finally, save
    doc.save(str(output_path))


# If you want a quick test (without GUI), you can do:
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python -m modernizer.writer <input.docx> <output.docx>")
        sys.exit(1)

    in_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2])

    parsed = parse_legacy_docx_by_sequence(in_path)
    write_new_doc(parsed, out_path)
    print(f"Wrote new file to {out_path}")
