#!/usr/bin/env python3
"""R03 – Missing Fonts & Images Checker  (word-dtp-preflight)"""
import sys, json, zipfile, re
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

def get_system_fonts():
    """Return set of installed font family names (lowercase)."""
    try:
        import subprocess
        out = subprocess.check_output(["fc-list", ":family"], text=True, stderr=subprocess.DEVNULL)
        fonts = set()
        for line in out.splitlines():
            for part in line.split(","):
                fonts.add(part.strip().lower())
        return fonts
    except Exception:
        return set()

def run(docx_path):
    results = {"check": "fonts_images", "missing_fonts": [], "missing_images": [],
               "referenced_fonts": [], "total_images": 0, "issues": 0}
    sys_fonts = get_system_fonts()

    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())

        # ── Fonts ─────────────────────────────────────────────────────────────
        referenced = set()
        for xmlfile in ["word/document.xml", "word/styles.xml", "word/fontTable.xml"]:
            if xmlfile not in names:
                continue
            root = etree.fromstring(zf.read(xmlfile))
            for el in root.iter(f"{{{W}}}rFonts"):
                for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia",
                             f"{{{W}}}ascii", f"{{{W}}}hAnsi", f"{{{W}}}cs", f"{{{W}}}eastAsia"):
                    v = el.get(attr)
                    if v:
                        referenced.add(v.strip())
            # fontTable entries
            for el in root.iter(f"{{{W}}}font"):
                v = el.get(f"{{{W}}}name") or el.get("w:name")
                if v:
                    referenced.add(v.strip())

        results["referenced_fonts"] = sorted(referenced)
        if sys_fonts:
            missing = [f for f in referenced if f.lower() not in sys_fonts
                       and not any(f.lower() in sf for sf in sys_fonts)]
            results["missing_fonts"] = missing
        else:
            results["missing_fonts"] = []  # Can't check without fc-list

        # ── Images ────────────────────────────────────────────────────────────
        media_files = {n for n in names if n.startswith("word/media/")}
        results["total_images"] = len(media_files)

        # Check rels for referenced images
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in names:
            root = etree.fromstring(zf.read(rels_path))
            for rel in root:
                rtype = rel.get("Type", "")
                target = rel.get("Target", "")
                if "image" in rtype.lower():
                    full = f"word/{target}" if not target.startswith("word/") else target
                    full = full.replace("word/../", "")
                    if full not in names and target not in names:
                        results["missing_images"].append(target)

    results["issues"] = len(results["missing_fonts"]) + len(results["missing_images"])
    return results

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: check_fonts_images.py <file.docx>"); sys.exit(1)
    print(json.dumps(run(sys.argv[1]), ensure_ascii=False, indent=2))
