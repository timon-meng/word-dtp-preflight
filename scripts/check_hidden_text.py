#!/usr/bin/env python3
"""R20 – Hidden Text Checker  (word-dtp-preflight)"""
import sys, json, zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def find_vanish_runs(root, source_label):
    found = []
    for i, p in enumerate(root.findall(f"{{{W}}}p", root.nsmap if hasattr(root, 'nsmap') else {})):
        pass
    for i, p in enumerate(root.findall(f".//{{{W}}}p")):
        for r in p.findall(f".//{{{W}}}r"):
            rpr = r.find(f"{{{W}}}rPr")
            if rpr is not None and rpr.find(f"{{{W}}}vanish") is not None:
                text = "".join(t.text or "" for t in r.findall(f"{{{W}}}t")).strip()
                if text:
                    found.append({
                        "source": source_label,
                        "paragraph": i + 1,
                        "text_preview": text[:80]
                    })
    return found

def run(docx_path):
    results = {
        "check": "hidden_text",
        "hidden_runs": [],
        "issues": 0,
    }

    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())
        targets = (["word/document.xml"] +
                   [n for n in names if n.startswith("word/header") or n.startswith("word/footer")] +
                   [n for n in names if "footnote" in n or "endnote" in n])

        for xml_path in targets:
            if xml_path not in names:
                continue
            root = etree.fromstring(zf.read(xml_path))
            label = xml_path.replace("word/", "").replace(".xml", "")
            results["hidden_runs"].extend(find_vanish_runs(root, label))

    results["issues"] = len(results["hidden_runs"])
    return results

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: check_hidden_text.py <file.docx>"); sys.exit(1)
    print(json.dumps(run(sys.argv[1]), ensure_ascii=False, indent=2))
