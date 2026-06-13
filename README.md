# Doc Manager Skill for Antigravity

Welcome to the Doc Manager skill. This tool lets your Antigravity agent create, edit, and manage Microsoft Word (.docx) documents on your local machine.

## How to install and set up

1. Folder placement: Put the entire `doc-manager` folder inside your Antigravity skills directory, typically `~/.gemini/config/skills/doc-manager`.
2. Prerequisites:
   - Python: Ensure Python is installed on your Windows machine.
   - uv: This skill uses `uv` to run Python scripts and handle dependencies. Install it in PowerShell with:
     `irm https://astral.sh/uv/install.ps1 | iex`
3. Once the folder is in place and `uv` is installed, Antigravity will detect the skill by reading the `SKILL.md` file. No additional setup is required.

## How to use

Ask Antigravity to perform the task you need; you do not need to run commands manually.

Creating documents
You can request full, professional documents with cover pages, tables of contents, callouts, and other sections.
Example prompts:
- "Create a professional Word document in my Downloads folder summarizing the Q3 marketing results. Include a cover page, an executive summary, and a timeline."
- "Generate a technical spec document for the new API. Use the 'technical' theme."

Editing existing documents
Antigravity can read and modify existing documents.
Example prompts:
- "Read `C:\Users\YourName\Downloads\Report.docx` and add a new section called 'Conclusion' at the end."
- "Replace the 'Timeline' section in my project Word document with this new schedule."
- "Find 'OldProjectName' in my doc and replace it with 'Project Phoenix'."

Features
- Themes: `professional` (default), `technical`, `executive`, `startup`.
- Block types: cover pages, headings, paragraphs, bullet/numbered lists, tables (with headers and shading), checklists, timelines, code blocks, executive summaries, and note/warning callout boxes.

Where files are saved
By default, generated or edited Word documents are saved to your Windows Downloads folder unless you specify a different path in your prompt.
