#!/usr/bin/env python3
"""
pack.py — Rebuild a .docx from a directory produced by unpack.py.

Part of the word-agent XML fallback chain. The output is written to a temp
file, verified as a readable OOXML package, then atomically moved into place.

Usage:
    python3 pack.py <src_dir> <output.docx>
"""

import argparse
import os
import sys
import tempfile
import zipfile


def pack(src_dir, output_path):
    if not os.path.isdir(src_dir):
        print(f"ERROR: not a directory: {src_dir}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(os.path.join(src_dir, "[Content_Types].xml")):
        print("ERROR: [Content_Types].xml missing — not an unpacked docx dir.",
              file=sys.stderr)
        sys.exit(1)

    out_dir = os.path.dirname(os.path.abspath(output_path)) or "."
    fd, tmp = tempfile.mkstemp(suffix=".docx", dir=out_dir)
    os.close(fd)
    count = 0
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            # [Content_Types].xml first is conventional for OOXML packages
            entries = []
            for base, _dirs, files in os.walk(src_dir):
                for name in files:
                    full = os.path.join(base, name)
                    rel = os.path.relpath(full, src_dir).replace(os.sep, "/")
                    entries.append((rel, full))
            entries.sort(key=lambda e: (e[0] != "[Content_Types].xml", e[0]))
            for rel, full in entries:
                zout.write(full, rel)
                count += 1
        with zipfile.ZipFile(tmp, "r") as zf:
            if "word/document.xml" not in zf.namelist():
                raise ValueError("word/document.xml missing from package")
        os.replace(tmp, output_path)
    except Exception:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise
    print(f"Packed {count} entries → {output_path}")


def main():
    ap = argparse.ArgumentParser(description="Rebuild a .docx from an unpacked directory")
    ap.add_argument("src", help="Directory produced by unpack.py")
    ap.add_argument("output", help="Output .docx path")
    args = ap.parse_args()
    pack(args.src, args.output)


if __name__ == "__main__":
    main()
