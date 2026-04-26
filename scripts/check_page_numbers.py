#!/usr/bin/env python3
"""R09 – Page Number Checker  (word-dtp-preflight)"""
import sys, json, zipfile, re
from lxml import etree

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

PAGE_INSTRS = re.compile(r"\b(PAGE|NUMPAGES|SECTIONPAGES)\b", re.I)

def has_page_field(xml_root):
    for instr in xml_root.iter(f"{{{W}}}instrText"):
        if instr.text and PAGE_INSTRS.search(instr.text):
            return True
    for fld in xml_root.iter(f"{{{W}}}fldSimple"):
        instr = fld.get(f"{{{W}}}instr") or ""
        if PAGE_INSTRS.search(instr):
            return True
    return False

def plain_text_numbers(xml_root):
    suspects = []
    for p in xml_root.findall(f".//{{{W}}}p"):
        text = "".join(t.text or "" for t in p.iter(f"{{{W}}}t")).strip()
        if re.fullmatch(r"\s*\d{1,4}\s*", text):
            suspects.append(text.strip())
    return suspects

def run(docx_path):
    results = {
        "check": "page_numbers",
        "sections_with_auto_pagenum": 0,
        "sections_without_pagenum": [],
        "plain_text_page_numbers": [],
        "issues": 0,
    }

    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())

        hf_files = [n for n in names if
                    (n.startswith("word/header") or n.startswith("word/footer"))
                    and n.endswith(".xml")]

        auto_count = 0
        plain_texts = []
        no_pagenum_files = []

        for hf in hf_files:
            root = etree.fromstring(zf.read(hf))
            if has_page_field(root):
                auto_count += 1
            else:
                text = "".join(t.text or "" for t in root.iter(f"{{{W}}}t")).strip()
                if text:
                    pt = plain_text_numbers(root)
                    if pt:
                        plain_texts.extend(pt)
                    else:
                        no_pagenum_files.append(hf)

        results["sections_with_auto_pagenum"] = auto_count
        results["plain_text_page_numbers"] = plain_texts
        results["sections_without_pagenum"] = no_pagenum_files

    results["issues"] = len(plain_texts) + len(no_pagenum_files)
    return results

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: check_page_numbers.py <file.docx>"); sys.exit(1)
    print(json.dumps(run(sys.argv[1]), ensure_ascii=False, indent=2))
