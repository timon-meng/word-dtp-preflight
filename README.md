# word-dtp-preflight

A **Claude Skill** for DTP (Desktop Publishing) teams — automatically runs quality checks on Microsoft Word (`.docx`) files as part of the Preflight and DTP production pipeline.

## What it does

Upload a `.docx` file to Claude and say *"帮我做预检"* or *"run preflight check"* — Claude will automatically run the selected checks and output a bilingual (Chinese/English) report.

### Supported Checks

| ID | Check | Description | Type |
|----|-------|-------------|------|
| R03 | Missing Fonts / Images | Detect referenced fonts not installed and broken image links | Analysis |
| R04 | Cross References | Find broken TOC entries, missing bookmarks, broken REF fields | Analysis |
| R05 | Extra Returns | Flag consecutive empty paragraphs and unnecessary soft returns | Analysis |
| R09 | Page Numbers | Detect hardcoded (plain-text) page numbers vs. auto field codes | Analysis |
| R20 | Hidden Text | Find text hidden via `w:vanish` style across body, headers, footers | Analysis |
| R21 | Numbering / Bullet Lists | Compare source and target file list structure and item count | Analysis |
| R29 | Numbers Comparison | Compare numeric values (integers, decimals, %, ranges) between source and target | Analysis |
| R25 | Unhide All Text | Remove all `w:vanish` properties — outputs a new file | **Write** |
| R16 | Pseudonymization | Replace all editable characters with `x` for anonymization — outputs a new file | **Write** |

## Installation

1. Download `word-dtp-preflight.skill`
2. Double-click to install in Claude Cowork
3. Or place the folder under your Claude skills directory

## Usage Examples

```
# Full preflight (runs R03–R20)
"帮我对这个 Word 文件做完整的预检"
"Run a full preflight check on this Word document"

# Specific check
"检查这个文件的交叉引用"
"Check cross references in this file"

# Comparison checks (needs source + target)
"比对这两个文件的编号列表和数字"
"Compare numbering and numbers between the source and translated file"

# Write operations
"取消隐藏这个文件里的所有文字"
"Pseudonymize this document for anonymization"
```

## Technical Details

- **Language**: Python 3 (pure standard library + `lxml`, `python-docx`)
- **Approach**: Directly parses DOCX ZIP/XML structure — no Microsoft Word installation required
- **Output**: Never modifies original files. Write operations produce `_cleaned` / `_unhidden` / `_pseudo` suffixed output files.
- **Dependencies**: `pip install python-docx lxml`

## Scripts

```
scripts/
├── check_fonts_images.py      # R03
├── check_cross_references.py  # R04
├── check_returns.py           # R05
├── check_page_numbers.py      # R09
├── check_hidden_text.py       # R20
├── check_numbering.py         # R21
├── check_numbers.py           # R29
├── unhide_text.py             # R25 (write)
└── pseudonymize.py            # R16 (write)
```

## Background

This skill was built for a DTP team's automation initiative, addressing requests from the `Automation_Requests_DTP` project. It replaces manual file-by-file checking with intelligent, conversational automation through Claude.

## License

MIT
