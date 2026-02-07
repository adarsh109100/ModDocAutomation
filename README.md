# JIRA Mod Doc Automation

Generate a Word Mod Document from Jira issues using a project-pinned template and a simple Flask UI. Works on Windows and Linux (project-relative paths only).

## Features
- Web UI with two tabs:
  - Generate Mod Doc: enter JQL and metadata (Version, Release, Revision, Build), create and download the doc.
  - JIRA API Configuration: enter Base URL, Email, API Token, save and test Jira connection.
- Splits modifications into two placeholders:
  - `<<JIRA_MODIFICATIONS>>` for Story issues
  - `<<JIRA_MODIFICATIONS_FORDEFECTANDIMP>>` for Defect/Improvement issues (Internal Defect, Defect, Internal Defect (MyPLAY), Improvement (MyPLAY), Improvement, Defect (migrated), Field Defect)
- Project-pinned template at `MyTemplates/MyTemplate.docx`.
- Comprehensive placeholder replacement across body, headers, footers, and tables.
- Output file saved under `static/` with timestamp suffix for portability.

## Project Structure
```
ModDocAutomation/
  app.py                    # Flask UI server
  JIraOLK.py                # Jira logic and document generation
  jira_config.json          # Saved Jira credentials (created/updated by the UI) [ignored by Git]
  jira_config.example.json  # Example config committed to repo (no secrets)
  MyTemplates/
    MyTemplate.docx         # Your Word template (placeholders described below)
  templates/
    index.html              # UI HTML
  static/
    css/style.css           # Optional custom styles
    JiraNotesSummary.docx   # Output base name (actual files include timestamp)
```

## Requirements
- Python 3.10+
- Dependencies:
  - Flask
  - python-docx
  - requests
  - python-dotenv

Install with:
```powershell
python -m pip install flask python-docx requests python-dotenv
```
On Linux/macOS:
```bash
python3 -m pip install flask python-docx requests python-dotenv
```

## Running the App
1) Start the Flask server:
```powershell
py app.py
```
On Linux/macOS:
```bash
python3 app.py
```
2) Open the UI in your browser at:
```
http://127.0.0.1:5000/
```

## Configure Jira
- Go to the "JIRA API Configuration" tab.
- Fill:
  - JIRA Base URL: `https://yourdomain.atlassian.net`
  - JIRA Email: your Atlassian email
  - JIRA API Token: a valid Jira Cloud API token
- Click "Save Configuration".
- Click "Test Jira Connection". You should see a message like:
  - "Connection successful. Authenticated as: <Your Name>"
  - Or a failure message with status code and response.

The configuration is stored in `jira_config.json`.

## Generate and Download Mod Doc
- Go to the "Generate Mod Doc" tab.
- Enter your JQL (e.g. `project = OLK AND fixVersion = 64766 ORDER BY created ASC`).
- Optionally fill metadata fields:
  - Onelink Version
  - Release
  - Revision
  - Build
- Click "Create & Download Mod Doc".
- A Word file will be generated based on your template and downloaded.
- A copy is saved under `static/` with a timestamp suffix.

## Template Placeholders
Use these exact tokens in `MyTemplates/MyTemplate.docx`:
- Jira content blocks:
  - `<<JIRA_MODIFICATIONS>>` → Story issues
  - `<<JIRA_MODIFICATIONS_FORDEFECTANDIMP>>` → Defect/Improvement issues
- Metadata:
  - `<<version>>` or `<<VERSION>>`
  - `<<Release>>` (also supports `<< Release>>` and `<< RELEASE >>`)
  - `<<Revision>>` (also supports `<< Revision >>` and `<< REVISION >>`)
  - `<<build>>` (also supports `<< Build >>` and `<<BUILD>>`)
- Date:
  - `<<DATE>>` (replaced everywhere, including headers/footers)
- Version (from Jira fixVersions or metadata):
  - `<<VERSION>>`

Notes:
- Do not use XML-like closing tags (e.g., `<<version>></version>`). Only the literal tokens above.
- Placeholders can be in body text, headers, footers, or table cells.

## How It Works
- `app.py` provides the UI and endpoints:
  - `/save-config` stores Jira credentials in `jira_config.json`.
  - `/test-connection` calls Jira `/rest/api/3/myself` with saved creds and shows a flash message.
  - `/generate-and-download` calls `generate_mod_doc(jql, meta)` and streams the generated doc.
- `JIraOLK.py` handles Jira calls and document generation:
  - Fetches issues via `/rest/api/3/search/jql` with your JQL.
  - Groups rich-text field content by headings and details.
  - Filters issues into Story vs Defect/Improvement.
  - Replaces placeholders across the entire document.
  - Saves output under `static/` with a timestamp.

## Troubleshooting
- "Connection failed" on Test Jira:
  - Verify Base URL, Email, and API Token are correct.
  - Base URL should be without a trailing slash (the app normalizes this).
- No content at `<<JIRA_MODIFICATIONS_FORDEFECTANDIMP>>`:
  - Confirm your JQL returns issues with one of the configured types:
    - Internal Defect, Defect, Internal Defect (MyPLAY), Improvement (MyPLAY), Improvement, Defect (migrated), Field Defect
  - Ensure `MODIFICATION_FIELD_ID` environment variable references the rich-text custom field ID containing the modifications.
- Placeholders not replaced:
  - Ensure the tokens match exactly (case-sensitive), and appear as one contiguous token (long placeholders can be split into multiple runs in Word—this code handles common splits).
- Date missing on first page header/footer:
  - Some templates have separate first/even page headers/footers; replacement attempts all variants, but confirm your template enables them if needed.

## Configuration via Environment (Optional)
You can override defaults with these environment variables:
- `OUTPUT_PATH` → Base output file name under `static/` (timestamp is added automatically)
- `MODIFICATION_FIELD_ID` → Jira custom field ID for modifications
- `JIRA_BASE_URL`, `JIRA_EMAIL`, `JIRA_API_TOKEN` → If not using UI config

Example (PowerShell):
```powershell
$env:JIRA_BASE_URL="https://yourdomain.atlassian.net"
$env:JIRA_EMAIL="you@example.com"
$env:JIRA_API_TOKEN="your_api_token"
$env:MODIFICATION_FIELD_ID="customfield_42223"  # replace with your field
$env:OUTPUT_PATH="static\JiraNotesSummary.docx"
py JIraOLK.py
```

## Hosting on Linux VM
- Clone/copy the project folder.
- Install dependencies with `python3 -m pip install -r requirements.txt` (or use the install commands above).
- Run `python3 app.py`.
- All paths are project-relative, so nothing references your local desktop.

## Security & Secrets
- This repo includes a `.gitignore` that excludes `jira_config.json` (contains tokens) and `.env`.
- Commit `jira_config.example.json` (no secrets) to share structure.
- If you already committed secrets earlier, rotate your Jira API token immediately and force-remove secrets from history.

## License
Internal use. Update this section if you plan to distribute.
