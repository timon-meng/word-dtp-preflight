# Check Scripts — Detailed Reference

## R03: check_fonts_images.py

**Logic:**
- Unpack DOCX as ZIP
- Parse `word/document.xml` and `word/fontTable.xml` for referenced font names
- Compare referenced fonts against system fonts (using `fc-list` on Linux / `font_manager` via matplotlib if available, fallback to a curated list)
- Parse `word/_rels/document.xml.rels` for image relationships
- Check each image file actually exists in the ZIP (`word/media/`)

**Output JSON:**
```json
{
  "check": "fonts_images",
  "missing_fonts": ["FontName1", "FontName2"],
  "missing_images": ["image1.png"],
  "embedded_fonts": ["EmbeddedFont"],
  "total_images": 5,
  "issues": 2
}
```

---

## R04: check_cross_references.py

**Logic:**
- Extract raw XML from `word/document.xml`
- Find all `<w:fldChar>` / `<w:instrText>` field instructions
- Flag fields containing "Error! Bookmark not defined" or "Error! Reference source not found"
- Find `<w:bookmarkStart>` elements; collect all bookmark names
- Find REF/PAGEREF field instructions; check if target bookmark exists
- Find hyperlinks (`<w:hyperlink>`); flag relative-path ones as potentially risky

**Output JSON:**
```json
{
  "check": "cross_references",
  "broken_fields": [{"text": "Error! Bookmark not defined", "context": "..."}],
  "missing_bookmarks": [{"ref": "BookmarkName", "paragraph": 12}],
  "toc_fields": 2,
  "index_fields": 0,
  "issues": 3
}
```

---

## R05: check_returns.py

**Logic:**
- Iterate all paragraphs via python-docx
- Hard returns (paragraph breaks): flag consecutive empty paragraphs (2+ in a row)
- Soft returns (line breaks): count `<w:br w:type="textWrapping"/>` elements; flag those not inside tables or figure captions
- Also check table cells for multiple trailing empty paragraphs

**Output JSON:**
```json
{
  "check": "returns",
  "extra_hard_returns": [{"location": "Para 7-8", "context": "between sections"}],
  "extra_soft_returns": [{"location": "Para 12", "context": "mid-sentence line break"}],
  "issues": 4
}
```

---

## R09: check_page_numbers.py

**Logic:**
- Extract headers/footers from `word/header*.xml` and `word/footer*.xml`
- Detect page number fields: `<w:fldSimple w:instr="PAGE">` or `<w:instrText>` containing PAGE
- Check if page numbers use field codes (automatic) vs plain text (manual/hardcoded)
- Count sections; if multiple sections, verify restart numbering settings
- Flag plain-text numbers that look like page numbers (digits in header/footer without field codes)

**Output JSON:**
```json
{
  "check": "page_numbers",
  "sections": 2,
  "auto_page_number_sections": 2,
  "manual_page_numbers": [],
  "plain_text_numbers_in_header_footer": ["Page 1", "Page 2"],
  "issues": 2
}
```

---

## R20: check_hidden_text.py

**Logic:**
- Scan all runs for `<w:vanish/>` property (hidden text style)
- Scan all text boxes / drawing objects: check if text content overflows the shape bounds (requires reading txbx geometry — use heuristic: if charCount * avg_char_width > box_width, flag)
- Check for `<w:rPr><w:vanish/></w:rPr>` in headers, footers, text boxes too

**Output JSON:**
```json
{
  "check": "hidden_text",
  "hidden_runs": [{"paragraph": 5, "text_preview": "This text is hidden..."}],
  "potentially_overflowing_textboxes": 0,
  "issues": 3
}
```

---

## R21: check_numbering.py

**Logic (requires source + target file):**
- Extract all list paragraphs from both files using python-docx (`paragraph.style.name` contains "List" or `paragraph.paragraph_format.numPr`)
- Normalize list sequences: extract numbering text, strip translations, compare structure
- Flag: extra items, missing items, different numbering style (bullet vs numbered)
- Compare total list paragraph counts

**Output JSON:**
```json
{
  "check": "numbering",
  "source_list_count": 12,
  "target_list_count": 11,
  "mismatches": [{"source": "3. Item three", "target": "missing", "location": "Para 45"}],
  "issues": 1
}
```

---

## R29: check_numbers.py

**Logic (requires source + target file):**
- Extract all numbers from both documents using regex: integers, decimals, percentages, ranges (e.g., "1-5"), measurements (e.g., "3.5 mm")
- Match numbers by proximity/paragraph position
- Flag numbers present in source but absent or changed in target
- Special handling for figure numbers ("Figure 3"), table numbers ("Table 1"), and data in tables

**Output JSON:**
```json
{
  "check": "numbers",
  "source_numbers": 45,
  "target_numbers": 44,
  "mismatches": [{"source_num": "3.5", "target_num": "35", "context": "Table 2, row 4"}],
  "issues": 1
}
```

---

## R25: unhide_text.py (WRITE OPERATION)

**Logic:**
- Load DOCX with python-docx
- Iterate all runs in paragraphs, headers, footers, text boxes
- Remove `<w:vanish/>` from run properties wherever found
- Save to output path (never overwrite original)

**Usage:** `python3 unhide_text.py input.docx output_unhidden.docx`

**Output JSON:**
```json
{
  "operation": "unhide_text",
  "runs_unhidden": 14,
  "output_file": "output_unhidden.docx"
}
```

---

## R16: pseudonymize.py (WRITE OPERATION)

**Logic:**
- Load DOCX, iterate all text runs in: body paragraphs, headers, footers, text boxes, table cells
- Replace every character that is NOT whitespace, newline, or punctuation with "x"
- Preserve: whitespace, punctuation, numbers (configurable), paragraph structure
- Option: replace with "x" keeping original character count; or use fixed word "XXXXX"
- Save to output path

**Usage:** `python3 pseudonymize.py input.docx output_pseudo.docx [--keep-numbers] [--keep-punctuation]`

**Output JSON:**
```json
{
  "operation": "pseudonymize",
  "runs_processed": 203,
  "chars_replaced": 8450,
  "output_file": "output_pseudo.docx"
}
```
