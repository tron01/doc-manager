# Doc Manager Skill for Antigravity

Welcome to the Doc Manager skill. This tool lets your Antigravity agent create, edit, and manage Microsoft Word (.docx) and PDF documents on your local machine.

## How to install and set up

1. Folder placement: Put the entire `doc-manager` folder inside your Antigravity skills directory.
   - Windows: `C:\Users\<YourName>\.gemini\config\skills\doc-manager`
   - macOS/Linux: `~/.gemini/config/skills/doc-manager`
2. Prerequisites:
   - Python: Ensure Python is installed on your machine.
   - uv: This skill uses `uv` to run Python scripts and handle dependencies.
     - Windows (PowerShell): `irm https://astral.sh/uv/install.ps1 | iex`
     - macOS/Linux: follow the `uv` installation instructions for your shell or package manager.
3. Once the folder is in place and `uv` is installed, Antigravity will detect the skill by reading the `SKILL.md` file. No additional setup is required.

## How to use

Ask Antigravity to perform the task you need; you do not need to run commands manually.

Creating documents
You can request full, professional documents with cover pages, tables of contents, callouts, images, charts, and other sections. You can also output as PDF.
Example prompts:
- "Create a professional Word document in my Downloads folder summarizing the Q3 marketing results. Include a cover page, an executive summary, and a timeline."
- "Generate a technical spec document for the new API. Use the 'technical' theme and export it as a PDF."
- "Batch generate 5 certificates from this CSV file."

Editing existing documents
Antigravity can read and modify existing documents.
Example prompts:
- "Read `C:\Users\YourName\Downloads\Report.docx` and add a new section called 'Conclusion' at the end."
- "Replace the 'Timeline' section in my project Word document with this new schedule."
- "Find 'OldProjectName' in my doc and replace it with 'Project Phoenix'."

Features
- Export Formats: `.docx` and `.pdf`
- Bulk Generation: Create multiple documents simultaneously from a CSV template.
- Themes: `professional` (default), `technical`, `executive`, `startup`.
- Block types: cover pages, headings, paragraphs, bullet/numbered lists, tables (with headers and shading), checklists, timelines, code blocks, executive summaries, note/warning callout boxes, images, and matplotlib charts.

Where files are saved
By default, generated or edited Word documents are saved to your Downloads folder unless you specify a different path in your prompt.
You can also provide an explicit path for Windows, macOS, or Linux in your request.
