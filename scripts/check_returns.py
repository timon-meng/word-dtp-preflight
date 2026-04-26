#!/usr/bin/env python3
"""R05 – Extra Returns Checker  (word-dtp-preflight)"""
import sys, json, zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def para_text(p):
    return "".join(t.text or "" for t in p.iter(f"{{{W}}}t")).strip()

def run(docx_path):
    results = {
        "check": "returns",
        "extra_hard_returns": [],
        "soft_returns": [],
        "issues": 0,
    }

    with zipfile.ZipFile(docx_path) as zf:
        if "word/document.xml" not in zf.namelist():
            return results
        root = etree.fromstring(zf.read("word/document.xml"))

    body = root.find(f"{{{W}}}body")
    if body is None:
        return results

    # ── Hard returns: consecutive empty paragraphs ────────────────────────────
    children = list(body)
    prev_empty = False
    for i, child in enumerate(children):
        tag = etree.QName(child.tag).localname if "{" in child.tag else child.tag
        if tag != "p":
            prev_empty = False
            continue
        text = para_text(child)
        is_empty = (text == "")
        if is_empty and prev_empty:
            results["extra_hard_returns"].append({
                "location": f"段落 {i+1}",
                "context": "连续空段落（多余硬回车）"
            })
        prev_empty = is_empty

    # ── Soft returns: <w:br> with type textWrapping ───────────────────────────
    for i, p in enumerate(root.findall(f".//{{{W}}}p")):
        parent = p.getparent()
        in_table = False
        el = parent
        while el is not None:
            ltag = etree.QName(el.tag).localname if "{" in el.tag else el.tag
            if ltag == "tbl":
                in_table = True
                break
            el = el.getparent()

        style_el = p.find(f".//{{{W}}}pStyle")
        style_name = (style_el.get(f"{{{W}}}val") or "") if style_el is not None else ""

        for br in p.findall(f".//{{{W}}}br"):
            br_type = br.get(f"{{{W}}}type") or br.get("w:type", "")
            if br_type in ("textWrapping", "") or br_type == "":
                if not in_table and "caption" not in style_name.lower():
                    ctx = para_text(p)[:60] or "（空段落）"
                    results["soft_returns"].append({
                        "location": f"段落 {i+1}",
                        "context": ctx
                    })
                    break

    results["issues"] = len(results["extra_hard_returns"]) + len(results["soft_returns"])
    return results

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: check_returns.py <file.docx>"); sys.exit(1)
    print(json.dumps(run(sys.argv[1]), ensure_ascii=False, indent=2))
