#!/usr/bin/env python3
"""R25 – Unhide All Text  (word-dtp-preflight)
Usage: unhide_text.py <input.docx> <output.docx>
"""
import sys, json, zipfile, shutil, tempfile, os
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def remove_vanish(root):
    count = 0
    for vanish in root.findall(f".//{{{W}}}vanish"):
        parent = vanish.getparent()
        if parent is not None:
            parent.remove(vanish)
            count += 1
    return count

def run(input_path, output_path):
    results = {"operation": "unhide_text", "runs_unhidden": 0, "output_file": output_path}

    with zipfile.ZipFile(input_path, "r") as src_zf:
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as dst_zf:
            for item in src_zf.infolist():
                data = src_zf.read(item.filename)
                if item.filename.endswith(".xml") and (
                    item.filename.startswith("word/") or "header" in item.filename or "footer" in item.filename
                ):
                    try:
                        root = etree.fromstring(data)
                        count = remove_vanish(root)
                        results["runs_unhidden"] += count
                        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                    except Exception:
                        pass
                dst_zf.writestr(item, data)

    return results

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: unhide_text.py <input.docx> <output.docx>"); sys.exit(1)
    result = run(sys.argv[1], sys.argv[2])
    print(json.dumps(result, ensure_ascii=False, indent=2))
