# modernizer/inventory.py

"""
inventory.py

Recursively scans all .docx under samples/legacy/,
uses parse_legacy_docx_with_keywords() to build a tree of:
  - document_title
  - sections (each with heading, content, children)
  - revision_history (table or written)
Then prints it with extra blank lines and two‐space indentation to match your example.
"""

import sys
import logging
from pathlib import Path
from typing import Dict, Any

# Adjust this import if you renamed the parser function
from .parser import parse_legacy_docx_by_sequence as parse_legacy_docx_with_subclauses



# Configure logging for troubleshooting
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def print_section_tree(node: Dict[str, Any], level: int = 0) -> None:
    """
    Recursively print a section node and its children, with two‐space indentation per level,
    and a blank line before each child block.
    """
    indent = "  " * level
    # Print this node’s heading and content‐line count
    print(f"{indent}- '{node['heading']}' (content lines: {len(node['content'])})")

    if node.get("children"):
        # Blank line before children
        print()
        for child in node["children"]:
            print_section_tree(child, level + 1)


def process_docx_file(docx_path: Path) -> None:
    """
    Parse one .docx and print the title, section tree, and revision history summary.
    """
    try:
        data = parse_legacy_docx_with_subclauses(docx_path)
    except Exception as e:
        logger.error(f"Failed to parse {docx_path.name}: {e}")
        return

    title = data.get("document_title", "")
    sections = data.get("sections", [])
    rev = data.get("revision_history", {})

    print(f"\n=== {docx_path.name} ===")
    print(f"Document Title: '{title}'")
    print(f"Top‐level sections found: {len(sections)}\n")

    for idx, section in enumerate(sections):
        print_section_tree(section, level=0)
        # Blank line between top‐level sections (but not after the last one)
        if idx < len(sections) - 1:
            print()

    # Finally, print revision‐history summary
    if rev.get("type") == "table":
        rows = rev.get("rows", [])
        print(f"\nRevision History: TABLE with {len(rows)} rows")
    else:
        content = rev.get("content", [])
        print(f"\nRevision History: WRITTEN with {len(content)} lines")


if __name__ == "__main__":
    # Locate project root and samples/legacy
    project_root = Path(__file__).parent.parent
    legacy_root = project_root / "samples" / "legacy"

    if not legacy_root.exists():
        print(f"Directory {legacy_root} does not exist.")
        sys.exit(1)

    # Find every .docx under samples/legacy, recursively
    docx_files = list(legacy_root.rglob("*.docx"))

    if not docx_files:
        print(f"No .docx files found under {legacy_root}")
        sys.exit(1)

    print(f"Found {len(docx_files)} .docx file(s) under {legacy_root}")
    for docx_file in docx_files:
        process_docx_file(docx_file)
