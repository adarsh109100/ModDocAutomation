import requests
from docx import Document
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import os
from datetime import datetime
from docx.table import Table

# === LOAD ENV VARIABLES ===
load_dotenv()

BASE_DIR = os.path.dirname(__file__)
# Pin template path to project
TEMPLATE_PATH = os.path.join(BASE_DIR, "MyTemplates", "MyTemplate.docx")
# OUTPUT pinned under project static dir; ensure directory exists
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)
OUTPUT_PATH = os.getenv("OUTPUT_PATH", os.path.join(STATIC_DIR, "JiraNotesSummary.docx"))

JIRA_BASE_URL = os.getenv("JIRA_BASE_URL")
JIRA_EMAIL = os.getenv("JIRA_EMAIL")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
MODIFICATION_FIELD_ID = os.getenv("MODIFICATION_FIELD_ID")

# === ISSUE TYPE GROUPS ===
# Story types for primary modifications placeholder
STORY_TYPES = {"Story"}
# Defect/Improvement types for secondary placeholder, provided by user
OTHERS_TYPES = {
    "Internal Defect",
    "Defect",
    "Internal Defect (MyPLAY)",
    "Improvement (MyPLAY)",
    "Improvement",  # include common name
    "Defect (migrated)",
    "Field Defect",
}

# === JQL QUERY ===
JQL_QUERY = 'fixVersion = 64766 AND project = OLK ORDER BY created ASC'

# === AUTHENTICATION ===
auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)
headers = {
    "Accept": "application/json"
}

# === RUNTIME CONFIG UPDATE ===
def set_jira_config(config: dict):
    global JIRA_BASE_URL, JIRA_EMAIL, JIRA_API_TOKEN, OUTPUT_PATH, auth
    JIRA_BASE_URL = config.get("JIRA_BASE_URL", JIRA_BASE_URL)
    JIRA_EMAIL = config.get("JIRA_EMAIL", JIRA_EMAIL)
    JIRA_API_TOKEN = config.get("JIRA_API_TOKEN", JIRA_API_TOKEN)
    OUTPUT_PATH = config.get("OUTPUT_PATH", OUTPUT_PATH)
    # TEMPLATE_PATH intentionally not overridden to keep it pinned to project
    auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)

# === Ensure default template exists ===
def ensure_default_template():
    # Ensure template dir exists
    os.makedirs(os.path.dirname(TEMPLATE_PATH), exist_ok=True)
    # Only create a minimal template if user's template is missing
    if not os.path.exists(TEMPLATE_PATH):
        doc = Document()
        section = doc.sections[0]
        header = section.header
        header_para = header.add_paragraph()
        header_para.add_run("Product Version: <<VERSION>>    Date: <<DATE>>")

        doc.add_paragraph("Modification Document")
        doc.add_paragraph("")
        doc.add_paragraph("Version: <<version>>")
        doc.add_paragraph("Release: <<Release>>")
        doc.add_paragraph("Revision: <<Revision>>")
        doc.add_paragraph("Build: <<build>>")
        doc.add_paragraph("")
        doc.add_paragraph("Changes:")
        doc.add_paragraph("<<JIRA_MODIFICATIONS>>")
        doc.add_paragraph("<<JIRA_MODIFICATIONS_FORDEFECTANDIMP>>")

        footer = section.footer
        footer_para = footer.add_paragraph()
        footer_para.add_run("Generated on <<DATE>>")

        doc.save(TEMPLATE_PATH)

# === FETCH ISSUES ===
def fetch_issues():
    url = f"{JIRA_BASE_URL.rstrip('/')}/rest/api/3/search/jql"
    params = {
        "jql": JQL_QUERY,
        "fields": MODIFICATION_FIELD_ID + ",fixVersions,issuetype",
        "maxResults": 100
    }
    try:
        response = requests.get(url, headers=headers, auth=auth, params=params, timeout=10)
        print(f"[DEBUG] Jira API status code: {response.status_code}")
        if response.status_code != 200:
            print(f"[DEBUG] Jira API response: {response.text}")
        response.raise_for_status()
        issues = response.json().get("issues", [])
        print(f"[DEBUG] Number of issues fetched: {len(issues)}")
        if issues:
            print(f"[DEBUG] Sample issue: {issues[0]}")
        return issues
    except requests.exceptions.RequestException as err:
        print(f"Network error: {err}")
        return []

# === GROUP ISSUES BY HEADING ===
def group_issues_by_heading(issuesdoc, field_id):
    grouped = {}
    for issue in issuesdoc:
        issue_key = issue["key"]
        field_data = issue["fields"].get(field_id)

        if not field_data:
            continue

        for block in field_data.get("content", []):
            if block["type"] == "paragraph":
                lines = []
                line = ""
                for inline in block.get("content", []):
                    if inline["type"] == "text":
                        line += inline["text"]
                    elif inline["type"] == "hardBreak":
                        lines.append(line.strip())
                        line = ""
                if line:
                    lines.append(line.strip())

                if not lines:
                    continue

                heading = lines[0]
                details = lines[1:] if len(lines) > 1 else ["No details provided"]

                if heading not in grouped:
                    grouped[heading] = []

                for detail in details:
                    grouped[heading].append((detail, issue_key))

    return grouped

# === REPLACE PLACEHOLDERS ===
def replace_placeholder_in_paragraphs(paragraphs, placeholder, new_text):
    for para in paragraphs:
        # Build full text from runs to handle split placeholders
        full_text = "".join(run.text for run in para.runs)
        if placeholder in full_text:
            new_full_text = full_text.replace(placeholder, new_text)
            # Clear existing runs content
            for run in para.runs:
                run.text = ""
            # Set paragraph text in a single run
            para.add_run(new_full_text)

def replace_placeholder_in_header_footer(doc, placeholder, new_text):
    for section in doc.sections:
        # Default header/footer
        replace_placeholder_in_paragraphs(section.header.paragraphs, placeholder, new_text)
        replace_placeholder_in_paragraphs(section.footer.paragraphs, placeholder, new_text)
        # First-page header/footer (when enabled in template)
        try:
            replace_placeholder_in_paragraphs(section.first_page_header.paragraphs, placeholder, new_text)
            replace_placeholder_in_paragraphs(section.first_page_footer.paragraphs, placeholder, new_text)
        except Exception:
            pass
        # Even-page header/footer (when enabled in template)
        try:
            replace_placeholder_in_paragraphs(section.even_page_header.paragraphs, placeholder, new_text)
            replace_placeholder_in_paragraphs(section.even_page_footer.paragraphs, placeholder, new_text)
        except Exception:
            pass

# === Comprehensive placeholder replacement ===
def replace_placeholders_everywhere(doc: Document, placeholders: dict):
    # Header/Footer
    for ph, val in placeholders.items():
        replace_placeholder_in_header_footer(doc, ph, val)

    # Body paragraphs
    for para in doc.paragraphs:
        for ph, val in placeholders.items():
            replace_placeholder_in_paragraphs([para], ph, val)

    # Tables (cells -> paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for ph, val in placeholders.items():
                        replace_placeholder_in_paragraphs([para], ph, val)

# === CREATE WORD DOCUMENT WITH TEMPLATE ===
def create_word_doc_with_template(issuesdoc, template_path, meta: dict | None = None):
    ensure_default_template()
    print(f"[DEBUG] Using template: {template_path}")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    doc = Document(template_path)

    # === Jira Fix Version ===
    fix_versions = []
    for issue in issuesdoc:
        versions = issue["fields"].get("fixVersions", [])
        for v in versions:
            fix_versions.append(v.get("name"))
    version_text = ", ".join(set(fix_versions)) if fix_versions else "vUnknown"

    today_date = datetime.now().strftime("%d-%b-%Y")

    # Replace placeholders for VERSION/DATE everywhere
    replace_placeholders_everywhere(doc, {"<<VERSION>>": version_text, "<<DATE>>": today_date})

    # New: replace metadata placeholders throughout the document
    if meta:
        variants = {
            "<<version>>": meta.get("version", ""),
            "<<VERSION>>": meta.get("version", ""),
            "<< Version >>": meta.get("version", ""),
            "<< Release>>": meta.get("release", ""),
            "<<Release>>": meta.get("release", ""),
            "<< RELEASE >>": meta.get("release", ""),
            "<< Revision >>": meta.get("revision", ""),
            "<<Revision>>": meta.get("revision", ""),
            "<< REVISION >>": meta.get("revision", ""),
            "<<build>>": meta.get("build", ""),
            "<< Build >>": meta.get("build", ""),
            "<<BUILD>>": meta.get("build", ""),
        }
        replace_placeholders_everywhere(doc, variants)

    # Replace placeholder in body for JIRA modifications
    grouped_all = group_issues_by_heading(issuesdoc, MODIFICATION_FIELD_ID)
    print(f"[DEBUG] Grouped data (all types): {grouped_all}")

    # Separate by issue type
    stories = filter_issues_by_type(issuesdoc, STORY_TYPES)
    others = filter_issues_by_type(issuesdoc, OTHERS_TYPES)
    # Debug counts
    print(f"[DEBUG] Story issues: {len(stories)} | Other(defect/impr) issues: {len(others)}")
    grouped_stories = group_issues_by_heading(stories, MODIFICATION_FIELD_ID)
    grouped_others = group_issues_by_heading(others, MODIFICATION_FIELD_ID)

    # Insert for stories at <<JIRA_MODIFICATIONS>>
    inserted_stories = insert_grouped_mods_at_placeholder(doc, "<<JIRA_MODIFICATIONS>>", grouped_stories)
    # Backward compatibility: if no stories placeholder found, fall back to all
    if not inserted_stories:
        insert_grouped_mods_at_placeholder(doc, "<<JIRA_MODIFICATIONS>>", grouped_all)

    # Insert for others at <<JIRA_MODIFICATIONS_FORDEFECTANDIMP>>
    insert_grouped_mods_at_placeholder(doc, "<<JIRA_MODIFICATIONS_FORDEFECTANDIMP>>", grouped_others)

    # Save with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(OUTPUT_PATH)
    # Ensure output directory exists
    out_dir = os.path.dirname(base)
    os.makedirs(out_dir, exist_ok=True)
    output_with_timestamp = f"{base}_{timestamp}{ext}"
    doc.save(output_with_timestamp)

    return output_with_timestamp

# Helper to filter issues by issuetype name
def filter_issues_by_type(issuesdoc, allowed_types: set):
    filtered = []
    for issue in issuesdoc:
        itype = (issue.get("fields", {}).get("issuetype", {}) or {}).get("name")
        if itype in allowed_types:
            filtered.append(issue)
    return filtered

# Insert grouped modifications at a given placeholder across paragraphs and tables
def insert_grouped_mods_at_placeholder(doc: Document, placeholder: str, grouped_data: dict):
    inserted = False
    empty_msg = "(No modifications found)" if not grouped_data else None
    # Body paragraphs
    for para in doc.paragraphs:
        if placeholder in ("".join(run.text for run in para.runs) or para.text):
            para.text = ""
            if empty_msg:
                para.add_run(empty_msg)
            else:
                for heading, items in grouped_data.items():
                    h = para
                    hr = h.add_run(("\n" if h.text else "") + heading)
                    hr.bold = True
                    hr.underline = True
                    for detail, issue_key in items:
                        para.add_run(f"\n• {detail} ({issue_key})")
            inserted = True
            break
    # Tables
    if not inserted:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        text = "".join(run.text for run in para.runs) or para.text
                        if placeholder in text:
                            para.text = ""
                            if empty_msg:
                                para.add_run(empty_msg)
                            else:
                                for heading, items in grouped_data.items():
                                    hr = para.add_run(("\n" if para.text else "") + heading)
                                    hr.bold = True
                                    hr.underline = True
                                    for detail, issue_key in items:
                                        para.add_run(f"\n• {detail} ({issue_key})")
                            inserted = True
                            break
                    if inserted:
                        break
                if inserted:
                    break
            if inserted:
                break
    return inserted

# === DIAGNOSTIC: Scan template placeholders ===
def scan_template_placeholders(path: str | None = None):
    """Load the template and print lines containing placeholder markers <<...>>.
    Scans headers, footers, body paragraphs, and tables.
    """
    tpl = path or TEMPLATE_PATH
    ensure_default_template()
    if not os.path.exists(tpl):
        print(f"Template not found at: {tpl}")
        return
    print(f"[SCAN] Reading template: {tpl}")
    doc = Document(tpl)

    def collect_from_paragraphs(paragraphs, scope: str):
        for i, para in enumerate(paragraphs):
            text = "".join(run.text for run in para.runs) or para.text
            if "<<" in text and ">>" in text:
                print(f"[{scope}] Paragraph {i}: {text}")

    # Headers/Footers
    for si, section in enumerate(doc.sections):
        collect_from_paragraphs(section.header.paragraphs, scope=f"Header S{si}")
        collect_from_paragraphs(section.footer.paragraphs, scope=f"Footer S{si}")

    # Body paragraphs
    collect_from_paragraphs(doc.paragraphs, scope="Body")

    # Tables
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                collect_from_paragraphs(cell.paragraphs, scope=f"Table{ti} R{ri}C{ci}")

# === GENERATE MOD DOC ===
def generate_mod_doc(jira_query, meta: dict | None = None):
    global JQL_QUERY
    JQL_QUERY = jira_query
    ensure_default_template()
    print(f"[DEBUG] Using template: {TEMPLATE_PATH}")
    issues = fetch_issues()
    output_file = create_word_doc_with_template(issues, TEMPLATE_PATH, meta)
    return output_file

# === MAIN EXECUTION ===
if __name__ == "__main__":
    try:
        # Diagnostic scan: verify template placeholders
        scan_template_placeholders()
        # Then try a generation using current config
        issues = fetch_issues()
        output_file = create_word_doc_with_template(issues, TEMPLATE_PATH)
        print(f" Document created: {output_file}")
    except Exception as e:
        print(f" Error: {e}")