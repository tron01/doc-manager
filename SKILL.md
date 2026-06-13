---
name: doc-manager
description: Create and edit Microsoft Word (.docx) and PDF documents locally. Supports batch generation, themes, cover pages, TOC, tables, images, charts, callouts, and code blocks. Can edit documents by inserting/replacing sections. Saves to the Windows Downloads folder by default. Use when the user asks to create a Word document or PDF, write a report, generate .docx files, or perform bulk document generation.
---

# Doc Manager v3.0 - Professional AI Document Generation

This skill uses a Python script (`scripts/docx_tool.py`) powered by `python-docx` to create and manipulate professional Word documents. Construct a JSON representation of the structure, and the tool builds the `.docx`.

> [!NOTE]
> Invoke commands using: `uv run --with python-docx scripts/docx_tool.py <command>`
> (Referred to as `TOOL` below for brevity).

## Themes

4 built-in themes handle fonts, colors, and styling:
* `professional` (Default) - Standard corporate style (blue accents)
* `technical` - Clean, dark accents, Segoe UI font
* `executive` - Elegant, large headings, deep navy accents
* `startup` - Modern, purple/indigo accents

## Commands

```bash
# 1. Validate content JSON before creating
TOOL validate --content content.json

# 2. Create a new document
TOOL create --title "Document_Title" --content content.json --theme professional --output-dir ~/Downloads

# 2b. Create a PDF instead of DOCX
TOOL create --title "Document_Title" --content content.json --output-format pdf

# 2c. Batch create documents from a template and CSV data
TOOL batch --template template.json --variables data.csv --output-format pdf

# 3. Get document info (useful before editing)
TOOL info --file ~/Downloads/Existing.docx --output info.json

# 4. Edit - Append blocks
TOOL edit --file ~/Downloads/Existing.docx --action append --content new_blocks.json

# 5. Edit - Insert after a specific heading
TOOL edit --file ~/Downloads/Existing.docx --action insert-after --heading "Introduction" --content insert_blocks.json

# 6. Edit - Replace an entire section (heading and its content)
TOOL edit --file ~/Downloads/Existing.docx --action replace-section --heading "Timeline" --content new_section.json

# 7. Edit - Delete a section
TOOL edit --file ~/Downloads/Existing.docx --action delete-section --heading "Appendix"

# 8. Edit - Find and Replace text
TOOL edit --file ~/Downloads/Existing.docx --action replace --find "OldName" --replace-with "NewName"

# 9. Edit - Update table by index
TOOL edit --file ~/Downloads/Existing.docx --action update-table --table-index 0 --table-data updated_table.json

# 10. Add/Remove Tables
TOOL add-table --file ~/Downloads/Existing.docx --table-data table.json
TOOL remove-table --file ~/Downloads/Existing.docx --index 1
```

## Content JSON Structure (v2)

To create or edit documents, save a JSON file in the following format and pass its path to the tool.

> [!WARNING]
> DO NOT pass complex JSON strings in the terminal. ALWAYS save to a file and pass `--content file.json`.

### Full Document Example

Note: Default values (`"bold": false`, `"italic": false`, `"theme": "professional"`) are omitted for brevity.

```json
{
  "metadata": {"title": "Quarterly Report", "author": "AI Assistant", "version": "1.0", "status": "Final", "department": "Engineering"},
  "header": {"text": "Project Phoenix"},
  "footer": {"page_numbers": true, "version": "1.0", "confidential": true},
  "content": [
    {"type": "cover_page", "title": "Quarterly Report", "subtitle": "Q3 Analysis", "author": "AI Assistant", "date": "October 2026"},
    {"type": "toc"},
    {"type": "executive_summary", "text": "This report summarizes Q3 performance."},
    {"type": "heading", "level": 1, "text": "Introduction"},
    {"type": "paragraph", "text": "Here is a standard paragraph."},
    {"type": "note", "title": "Information", "text": "This is a blue callout box useful for general information."},
    {"type": "code", "language": "python", "text": "def hello_world():\n    print('Hello')\n"},
    {"type": "checklist", "items": ["[x] Setup database", "[ ] Configure API"]},
    {"type": "timeline", "items": [{"phase": "Phase 1", "date": "Q1 2026", "description": "Design"}]},
    {"type": "table", "headers": ["Metric", "Value", "Trend"], "rows": [["Latency", "45ms", "Down"], ["Uptime", "99.9%", "Stable"]]},
    {"type": "page_break"}
  ]
}
```

### Supported Block Types

| Type | Required Keys | Notes |
|---|---|---|
| `cover_page` | `title` | Centered title page (forces page break). Best as first block. |
| `toc` | - | Table of Contents placeholder (requires user to update field in Word). |
| `heading` | `level`, `text` | Levels 1-4. |
| `paragraph` | `text` | Basic body text. Supports `bold`, `italic`. |
| `bullet_list` / `numbered_list` | `items` | Array of strings. |
| `table` | `headers`, `rows` | Themed header colors and alternating row shading. |
| `checklist` | `items` | Array of strings. Prefix with `[x]` or `[ ]` for checkboxes. |
| `timeline` | `items` | Specialized milestone table (`phase`, `date`, `description`). |
| `code` | `text` | Monospace font, shaded background. Optional `language`. |
| `executive_summary` | `text` | Blue-bordered shaded summary box. |
| `note` / `warning` / `important` | `text` | Callout boxes with icons. Optional `title`. |
| `page_break` | - | Forces a page break. |
| `image` | `path` | Embeds an image. Optional `width` (in inches), `alignment` (`center`\|`right`), `caption`. |
| `chart` | `chart_type`, `data` | Embeds a matplotlib chart. Types: `bar`, `pie`, `line`. `data` must contain `labels` and `values` or `datasets`. Optional `title`. |

## Template Substitution

For batch generation or single creation, `{{variable_name}}` syntax can be used in your JSON templates. The variables will be resolved using the `--variables` flag.

## Editing Existing Documents

Use the `info` command to see document structure. For edit actions (`append`, `insert-after`, `replace-section`), the JSON uses the same `content` array format as above.

## Best Practices

1. **Use files for JSON.** Never pass complex JSON strings inline.
2. **Validate first.** Use `validate --content file.json` before `create`.
3. **Cover Page placement.** Should be the first block in your `content` array.
4. **TOC needs manual update.** The user must press Ctrl+A, F9 in Word to populate page numbers.
5. **Output path.** Defaults to `~/Downloads`. Use `--output-dir` to override.
