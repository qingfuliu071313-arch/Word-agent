#!/usr/bin/env python3
"""
unpack.py — Extract a .docx (OOXML zip) into a directory for XML-level editing.

Part of the word-agent XML fallback chain (used when MCP backends are
unavailable). Pair with pack.py to rebuild the .docx afterwards.

Usage:
    python3 unpack.py <input.docx> <dest_dir> [--pretty]

--pretty pretty-prints .xml/.rels parts for easier line-based editing.
WARNING: pretty-printing inserts whitespace into the XML tree; w:t elements
with xml:space="preserve" are safe, but avoid --pretty if you plan to pack
the result for production use without reviewing whitespace-sensitive text.
"""

import argparse
import os
import sys
import zipfile


def pretty_xml(data):
    """Pretty-print XML bytes; returns original bytes on any parse failure."""
    try:
        import xml.dom.minidom as minidom
        dom = minidom.parseString(data)
        out = dom.toprettyxml(indent="  ", encoding="UTF-8")
        # minidom adds blank lines for existing whitespace nodes; drop them
        lines = [l for l in out.split(b"\n") if l.strip()]
        return b"\n".join(lines) + b"\n"
    except Exception:
        return data


def unpack(docx_path, dest_dir, pretty=False):
    if not zipfile.is_zipfile(docx_path):
        print(f"ERROR: not a valid .docx (zip) file: {docx_path}", file=sys.stderr)
        sys.exit(1)
    os.makedirs(dest_dir, exist_ok=True)
    count = 0
    with zipfile.ZipFile(docx_path, "r") as z:
        for entry in z.namelist():
            # Defend against path traversal in malformed archives
            target = os.path.normpath(os.path.join(dest_dir, entry))
            if not target.startswith(os.path.normpath(dest_dir) + os.sep):
                print(f"  SKIPPED unsafe entry: {entry}", file=sys.stderr)
                continue
            if entry.endswith("/"):
                os.makedirs(target, exist_ok=True)
                continue
            os.makedirs(os.path.dirname(target), exist_ok=True)
            data = z.read(entry)
            if pretty and (entry.endswith(".xml") or entry.endswith(".rels")):
                data = pretty_xml(data)
            with open(target, "wb") as f:
                f.write(data)
            count += 1
    print(f"Unpacked {count} entries → {dest_dir}")


def main():
    ap = argparse.ArgumentParser(description="Extract a .docx into a directory")
    ap.add_argument("docx", help="Input .docx file")
    ap.add_argument("dest", help="Destination directory")
    ap.add_argument("--pretty", action="store_true",
                    help="Pretty-print XML parts (see header warning)")
    args = ap.parse_args()
    if not os.path.exists(args.docx):
        print(f"ERROR: file not found: {args.docx}", file=sys.stderr)
        sys.exit(1)
    unpack(args.docx, args.dest, pretty=args.pretty)


if __name__ == "__main__":
    main()
