#!/usr/bin/env python3
"""R21 – Numbering / Bullet List Checker  (word-dtp-preflight)
Usage: check_numbering.py <source.docx> <target.docx>
"""
import sys, json, zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def extract_list_paragraphs(docx_path):
    items = []
    with zipfile.ZipFile(docx_path) as zf:
        if "word/document.xml" not in zf.namelist():
            return items
        root = etree.fromstring(zf.read("word/document.xml"))

    for i, p in enumerate(root.findall(f".//{{{W}}}p")):
        ppr = p.find(f"{{{W}}}pPr")
        if ppr is None:
            continue
        num_pr = ppr.find(f"{{{W}}}numPr")
        if num_pr is None:
            pstyle = ppr.find(f"{{{W}}}pStyle")
            style_val = (pstyle.get(f"{{{W}}}val") or "") if pstyle is not None else ""
            if not any(kw in style_val.lower() for kw in ["list", "bullet", "listparagraph"]):
                continue
        text = "".join(t.text or "" for t in p.iter(f"{{{W}}}t")).strip()
        ilvl_el = num_pr.find(f"{{{W}}}ilvl") if num_pr is not None else None
        ilvl = int(ilvl_el.get(f"{{{W}}}val", "0")) if ilvl_el is not None else 0
        items.append({"index": i, "level": ilvl, "text": text})
    return items

def run(source_path, target_path):
    results = {
        "check": "numbering",
        "source_list_count": 0,
        "target_list_count": 0,
        "level_mismatch": False,
        "count_difference": 0,
        "mismatches": [],
        "issues": 0,
    }

    src_items = extract_list_paragraphs(source_path)
    tgt_items = extract_list_paragraphs(target_path)
    results["source_list_count"] = len(src_items)
    results["target_list_count"] = len(tgt_items)
    results["count_difference"] = abs(len(src_items) - len(tgt_items))

    src_levels = [x["level"] for x in src_items]
    tgt_levels = [x["level"] for x in tgt_items]
    if src_levels != tgt_levels:
        results["level_mismatch"] = True

    for i, (s, t) in enumerate(zip(src_items, tgt_items)):
        if s["level"] != t["level"]:
            results["mismatches"].append({
                "item": i + 1,
                "source_level": s["level"], "source_text": s["text"][:60],
                "target_level": t["level"], "target_text": t["text"][:60],
                "issue": "缩进级别不一致"
            })

    if len(src_items) > len(tgt_items):
        for j in range(len(tgt_items), len(src_items)):
            results["mismatches"].append({
                "item": j + 1,
                "source_text": src_items[j]["text"][:60],
                "issue": "译文中缺少对应列表项"
            })

    results["issues"] = len(results["mismatches"])
    if results["level_mismatch"] and results["issues"] == 0:
        results["issues"] = 1
    return results

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: check_numbering.py <source.docx> <target.docx>"); sys.exit(1)
    print(json.dumps(run(sys.argv[1], sys.argv[2]), ensure_ascii=False, indent=2))
