#!/usr/bin/env python3
"""R16 – Pseudonymization (replace editable text with 'x')  (word-dtp-preflight)
Usage: pseudonymize.py <input.docx> <output.docx> [--keep-numbers] [--keep-punctuation]

Replaces every letter character with 'x', preserving:
  - Whitespace and newlines
  - Numbers (if --keep-numbers)
  - Punctuation (if --keep-punctuation, default: keep)
  - Paragraph/run structure
"""
import sys, json, zipfile, re
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def pseudonymize_text(text, keep_numbers=True, keep_punctuation=True):
    result = []
    for ch in text:
        if ch.isspace():
            result.append(ch)
        elif keep_numbers and ch.isdigit():
            result.append(ch)
        elif keep_punctuation and not ch.isalpha():
            result.append(ch)
        elif ch.isalpha():
            result.append('x')
        else:
            result.append('x')
    return "".join(result)

def process_xml(data, keep_numbers, keep_punctuation):
    count = [0]
    try:
        root = etree.fromstring(data)
    except Exception:
        return data, 0

    for t_el in root.findall(f".//{{{W}}}t"):
        if t_el.text:
            new_text = pseudonymize_text(t_el.text, keep_numbers, keep_punctuation)
            if new_text != t_el.text:
                count[0] += 1
            t_el.text = new_text

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), count[0]

def run(input_path, output_path, keep_numbers=True, keep_punctuation=True):
    results = {
        "operation": "pseudonymize",
        "runs_processed": 0,
        "chars_replaced": 0,
        "output_file": output_path
    }

    with zipfile.ZipFile(input_path, "r") as src_zf:
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as dst_zf:
            for item in src_zf.infolist():
                data = src_zf.read(item.filename)
                if item.filename.endswith(".xml") and item.filename.startswith("word/"):
                    new_data, c = process_xml(data, keep_numbers, keep_punctuation)
                    results["runs_processed"] += c
                    data = new_data
                dst_zf.writestr(item, data)

    return results

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: pseudonymize.py <input.docx> <output.docx> [--keep-numbers] [--keep-punctuation]")
        sys.exit(1)
    keep_n = "--keep-numbers" in sys.argv
    keep_p = "--keep-punctuation" in sys.argv or "--keep-numbers" not in sys.argv
    result = run(sys.argv[1], sys.argv[2], keep_n, keep_p)
    print(json.dumps(result, ensure_ascii=False, indent=2))
