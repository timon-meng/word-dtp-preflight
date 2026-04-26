#!/usr/bin/env python3
"""R04 – Cross Reference Checker  (word-dtp-preflight)"""
import sys, json, zipfile, re
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def iter_instr_text(root):
    """Yield all fldChar instruction text blocks."""
    current_instr = []
    in_field = False
    for el in root.iter():
        tag = etree.QName(el.tag).localname if "{" in el.tag else el.tag
        if tag == "fldChar":
            ftype = el.get(f"{{{W}}}fldCharType") or el.get("w:fldCharType", "")
            if ftype == "begin":
                in_field = True
                current_instr = []
            elif ftype == "end":
                if current_instr:
                    yield " ".join(current_instr).strip()
                in_field = False
        elif tag == "instrText" and in_field:
            current_instr.append(el.text or "")

def run(docx_path):
    results = {
        "check": "cross_references",
        "broken_fields": [],
        "missing_bookmarks": [],
        "toc_fields": 0, "index_fields": 0,
        "hyperlink_risks": [],
        "issues": 0,
    }

    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())
        if "word/document.xml" not in names:
            return results

        root = etree.fromstring(zf.read("word/document.xml"))

        # ── Collect bookmarks ─────────────────────────────────────────────────
        defined_bookmarks = set()
        for bm in root.iter(f"{{{W}}}bookmarkStart"):
            name = bm.get(f"{{{W}}}name") or bm.get("w:name", "")
            if name:
                defined_bookmarks.add(name)

        # ── Scan paragraphs for field errors ──────────────────────────────────
        paragraphs = root.findall(f".//{{{W}}}p")
        for i, para in enumerate(paragraphs):
            full_text = "".join(t.text or "" for t in para.iter(f"{{{W}}}t"))
            if "Error!" in full_text:
                results["broken_fields"].append({
                    "paragraph": i + 1,
                    "text_preview": full_text[:120]
                })

        # ── Scan field instructions ───────────────────────────────────────────
        for instr in iter_instr_text(root):
            instr_up = instr.upper().strip()
            if instr_up.startswith("TOC"):
                results["toc_fields"] += 1
            elif instr_up.startswith("INDEX"):
                results["index_fields"] += 1
            elif instr_up.startswith(("REF ", "PAGEREF ")):
                parts = instr.split()
                if len(parts) >= 2:
                    bm_name = parts[1].strip('"')
                    if bm_name not in defined_bookmarks:
                        results["missing_bookmarks"].append({
                            "instruction": instr[:80],
                            "bookmark": bm_name
                        })

        # ── Hyperlinks – flag relative paths ─────────────────────────────────
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in names:
            rel_root = etree.fromstring(zf.read(rels_path))
            for rel in rel_root:
                rtype = rel.get("Type", "")
                target = rel.get("Target", "")
                if "hyperlink" in rtype.lower():
                    if target and not target.startswith(("http", "mailto", "#")):
                        results["hyperlink_risks"].append(target)

    results["issues"] = (len(results["broken_fields"]) +
                         len(results["missing_bookmarks"]) +
                         len(results["hyperlink_risks"]))
    return results

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: check_cross_references.py <file.docx>"); sys.exit(1)
    print(json.dumps(run(sys.argv[1]), ensure_ascii=False, indent=2))
