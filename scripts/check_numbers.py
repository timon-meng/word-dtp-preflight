#!/usr/bin/env python3
"""R29 – Numbers Comparison Checker  (word-dtp-preflight)
Usage: check_numbers.py <source.docx> <target.docx>
"""
import sys, json, zipfile, re
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

NUM_RE = re.compile(
    r"""(?<![a-zA-Z])
        (
          \d{1,3}(?:,\d{3})*
          (?:\.\d+)?
          |\d+\.\d+
          |\d+
        )
        (?:\s*[-\u2013]\s*\d+)?
        (?:\s*%)?
    """,
    re.VERBOSE,
)

def extract_text_blocks(docx_path):
    blocks = []
    with zipfile.ZipFile(docx_path) as zf:
        if "word/document.xml" not in zf.namelist():
            return blocks
        root = etree.fromstring(zf.read("word/document.xml"))

    for i, p in enumerate(root.findall(f".//{{{W}}}p")):
        text = "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))
        if text.strip():
            blocks.append((f"段落{i+1}", text))
    return blocks

def normalize_num(s):
    s = s.replace(",", "")
    try:
        return str(float(s)) if "." in s else s
    except ValueError:
        return s

def run(source_path, target_path):
    results = {
        "check": "numbers",
        "source_number_count": 0,
        "target_number_count": 0,
        "mismatches": [],
        "issues": 0,
    }

    src_blocks = extract_text_blocks(source_path)
    tgt_blocks = extract_text_blocks(target_path)

    src_nums = []
    for loc, text in src_blocks:
        for m in NUM_RE.finditer(text):
            src_nums.append({"num": m.group(), "norm": normalize_num(m.group()), "loc": loc, "context": text[max(0,m.start()-20):m.end()+20]})

    tgt_nums = []
    for loc, text in tgt_blocks:
        for m in NUM_RE.finditer(text):
            tgt_nums.append({"num": m.group(), "norm": normalize_num(m.group()), "loc": loc})

    results["source_number_count"] = len(src_nums)
    results["target_number_count"] = len(tgt_nums)

    tgt_norm_list = [x["norm"] for x in tgt_nums]
    used_tgt = list(tgt_norm_list)
    for s in src_nums:
        if s["norm"] in used_tgt:
            used_tgt.remove(s["norm"])
        else:
            results["mismatches"].append({
                "source_number": s["num"],
                "location": s["loc"],
                "context": s["context"].strip()[:80],
                "issue": "译文中未找到对应数字"
            })

    results["issues"] = len(results["mismatches"])
    return results

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: check_numbers.py <source.docx> <target.docx>"); sys.exit(1)
    print(json.dumps(run(sys.argv[1], sys.argv[2]), ensure_ascii=False, indent=2))
