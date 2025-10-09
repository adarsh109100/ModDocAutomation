import requests
from docx import Document
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import os
from datetime import datetime

# === LOAD ENV VARIABLES ===
load_dotenv()

JIRA_BASE_URL = os.getenv("JIRA_BASE_URL")
JIRA_EMAIL = os.getenv("JIRA_EMAIL")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
MODIFICATION_FIELD_ID = os.getenv("MODIFICATION_FIELD_ID")
OUTPUT_PATH = os.getenv("OUTPUT_PATH", "C:/Users/AP109100/Desktop/pytest/JiraNotesSummary.docx")
TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "C:/Users/AP109100/Desktop/pytest/MyTemplate.docx")

# === JQL QUERY ===
JQL_QUERY = 'fixVersion = 64766 AND project = OLK ORDER BY created ASC'

# === AUTHENTICATION ===
auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)
headers = {
    "Accept": "application/json"
}

# === FETCH ISSUES ===
def fetch_issues():
    url = f"{JIRA_BASE_URL}/rest/api/3/search"
    params = {
        "jql": JQL_QUERY,
        "fields": MODIFICATION_FIELD_ID + ",fixVersions",
        "maxResults": 100
    }
    try:
        response = requests.get(url, headers=headers, auth=auth, params=params, timeout=10)
        response.raise_for_status()
        return response.json().get("issues", [])
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
        if placeholder in para.text:
            inline = para.runs
            for run in inline:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, new_text)

def replace_placeholder_in_header_footer(doc, placeholder, new_text):
    for section in doc.sections:
        header = section.header
        replace_placeholder_in_paragraphs(header.paragraphs, placeholder, new_text)

        footer = section.footer
        replace_placeholder_in_paragraphs(footer.paragraphs, placeholder, new_text)

# === CREATE WORD DOCUMENT WITH TEMPLATE ===
def create_word_doc_with_template(issuesdoc, template_path):
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

    # Replace placeholders in header/footer
    replace_placeholder_in_header_footer(doc, "<<VERSION>>", version_text)
    replace_placeholder_in_header_footer(doc, "<<DATE>>", today_date)

    # Replace placeholder in body
    grouped_data = group_issues_by_heading(issuesdoc, MODIFICATION_FIELD_ID)
    for para in doc.paragraphs:
        if "<<JIRA_MODIFICATIONS>>" in para.text:
            para.text = ""
            for heading, items in grouped_data.items():
                h = doc.add_paragraph()
                run = h.add_run(heading)
                run.bold = True
                run.underline = True

                for detail, issue_key in items:
                    bullet = doc.add_paragraph(style="List Bullet")
                    bullet.add_run(f"{detail} ({issue_key})")
            break

    # Save with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(OUTPUT_PATH)
    output_with_timestamp = f"{base}_{timestamp}{ext}"
    doc.save(output_with_timestamp)

    return output_with_timestamp

# === GENERATE MOD DOC ===
def generate_mod_doc(jira_query):
    global JQL_QUERY
    JQL_QUERY = jira_query
    issues = fetch_issues()
    output_file = create_word_doc_with_template(issues, TEMPLATE_PATH)
    return output_file

# === MAIN EXECUTION ===
if __name__ == "__main__":
    try:
        issues = fetch_issues()
        output_file = create_word_doc_with_template(issues, TEMPLATE_PATH)
        print(f" Document created: {output_file}")
    except Exception as e:
        print(f" Error: {e}")
