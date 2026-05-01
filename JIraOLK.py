import requests
from docx import Document
from requests.auth import HTTPBasicAuth
import os
from dotenv import load_dotenv
from datetime import datetime
from docx.table import Table

BASE_DIR = os.path.dirname(__file__)
# === LOAD ENV VARIABLES ===
DOTENV_PATH = os.path.join(BASE_DIR, ".env")
load_dotenv(dotenv_path=DOTENV_PATH, override=True)

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
LOG_LEVEL = os.getenv("LOG_LEVEL", os.getenv("log_level", "INFO")).upper()
OUTPUT_RETENTION_DAYS = int(os.getenv("OUTPUT_RETENTION_DAYS", "30"))
JIRA_PAGE_SIZE = int(os.getenv("JIRA_PAGE_SIZE", "100"))

LOG_LEVEL_ORDER = {
    "DEBUG": 10,
    "INFO": 20,
    "WARN": 30,
    "ERROR": 40,
}

LAST_RUN_SUMMARY = {}

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
    "ATF Defect",
    "Internal Defect - Art",
    "ATF/Regulator Defect",
}

# === JQL QUERY ===
JQL_QUERY = 'fixVersion = 64766 AND project = OLK ORDER BY created ASC'

# === AUTHENTICATION ===
auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)
headers = {
    "Accept": "application/json"
}

def _log(level: str, message: str):
    configured = LOG_LEVEL_ORDER.get(LOG_LEVEL, 20)
    current = LOG_LEVEL_ORDER.get(level, 20)
    if current >= configured:
        print(f"[{level}] {message}", flush=True)

if LOG_LEVEL not in LOG_LEVEL_ORDER:
    print(f"[WARN] Invalid LOG_LEVEL '{LOG_LEVEL}'. Falling back to INFO.", flush=True)
    LOG_LEVEL = "INFO"

print(f"[INFO] Logger initialized. level={LOG_LEVEL}, dotenv={DOTENV_PATH}", flush=True)

def get_last_run_summary():
    return dict(LAST_RUN_SUMMARY)

# === RUNTIME CONFIG UPDATE ===
def set_jira_config(config: dict):
    global JIRA_BASE_URL, JIRA_EMAIL, JIRA_API_TOKEN, OUTPUT_PATH, auth
    JIRA_BASE_URL = config.get("JIRA_BASE_URL", JIRA_BASE_URL)
    JIRA_EMAIL = config.get("JIRA_EMAIL", JIRA_EMAIL)
    JIRA_API_TOKEN = config.get("JIRA_API_TOKEN", JIRA_API_TOKEN)
    OUTPUT_PATH = config.get("OUTPUT_PATH", OUTPUT_PATH)
    # TEMPLATE_PATH intentionally not overridden to keep it pinned to project
    auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)

def _extract_text_from_adf_node(node):
    """Recursively extract readable text from Jira ADF nodes."""
    if isinstance(node, list):
        return "".join(_extract_text_from_adf_node(item) for item in node)
    if not isinstance(node, dict):
        return ""

    node_type = node.get("type")
    if node_type == "text":
        return node.get("text", "")
    if node_type == "hardBreak":
        return "\n"
    if node_type == "mention":
        attrs = node.get("attrs", {}) or {}
        return attrs.get("text") or attrs.get("displayName") or ""
    if node_type == "emoji":
        attrs = node.get("attrs", {}) or {}
        return attrs.get("text") or attrs.get("shortName") or ""
    if node_type == "inlineCard":
        attrs = node.get("attrs", {}) or {}
        return attrs.get("url", "")

    content = node.get("content", [])
    text = "".join(_extract_text_from_adf_node(child) for child in content)

    if node_type in {"paragraph", "heading", "listItem"} and text and not text.endswith("\n"):
        return text + "\n"
    return text

def _extract_lines_from_adf_node(node):
    text = _extract_text_from_adf_node(node)
    return [line.strip() for line in text.splitlines() if line.strip()]

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
    base = JIRA_BASE_URL.rstrip('/')
    urls = [
        f"{base}/rest/api/3/search/jql",
        f"{base}/rest/api/3/search",
    ]
    requested_fields = f"{MODIFICATION_FIELD_ID},fixVersions,issuetype"
    last_error = None

    for url in urls:
        all_issues = []
        start_at = 0
        total = None
        try:
            _log("INFO", f"Trying Jira search endpoint: {url}")
            while True:
                params = {
                    "jql": JQL_QUERY,
                    "fields": requested_fields,
                    "maxResults": JIRA_PAGE_SIZE,
                    "startAt": start_at,
                }
                response = requests.get(url, headers=headers, auth=auth, params=params, timeout=10)
                _log("DEBUG", f"Jira API page status code: {response.status_code} startAt={start_at}")
                if response.status_code != 200:
                    _log("WARN", f"Jira API request to {url} failed with status {response.status_code}")
                response.raise_for_status()

                payload = response.json()
                page_issues = payload.get("issues", [])
                page_total = payload.get("total")
                page_start = payload.get("startAt", start_at)
                page_max = payload.get("maxResults", JIRA_PAGE_SIZE)

                if total is None and page_total is not None:
                    total = page_total

                all_issues.extend(page_issues)
                _log(
                    "INFO",
                    f"Fetched Jira page: startAt={page_start}, pageSize={len(page_issues)}, "
                    f"accumulated={len(all_issues)}, total={total if total is not None else 'unknown'}",
                )

                if not page_issues:
                    break
                if total is not None and len(all_issues) >= total:
                    break
                if len(page_issues) < page_max:
                    break

                start_at = page_start + page_max

            _log("INFO", f"Total Jira issues fetched: {len(all_issues)}")
            return all_issues
        except requests.exceptions.RequestException as err:
            last_error = err
            _log("WARN", f"Jira search endpoint failed ({url}): {err}")

    _log("ERROR", f"Network error while fetching issues: {last_error}")
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
            lines = _extract_lines_from_adf_node(block)
            if not lines:
                continue

            heading = lines[0]
            details = lines[1:] if len(lines) > 1 else ["No details provided"]

            if heading not in grouped:
                grouped[heading] = []

            for detail in details:
                grouped[heading].append((detail, issue_key))

    return grouped

# === EMPTY MODIFICATION FIELD CHECK ===
def _collect_modification_text(field_data):
    """Extract plain text fragments from Jira rich-text modification field."""
    texts = []
    if not isinstance(field_data, dict):
        return texts

    for block in field_data.get("content", []):
        texts.extend(_extract_lines_from_adf_node(block))

    return texts

def is_modification_field_empty(field_data):
    """Return True when modification field has no usable text or only placeholder text."""
    texts = _collect_modification_text(field_data)
    if not texts:
        return True

    # Treat common placeholders as empty values.
    placeholders = {"tbd"}
    normalized = [t.strip().lower() for t in texts if t and t.strip()]
    if not normalized:
        return True
    return all(text in placeholders for text in normalized)

def print_empty_modification_ticket_keys(issuesdoc, field_id):
    """Print only Jira issue keys where modification field is empty."""
    for issue in issuesdoc:
        field_data = issue.get("fields", {}).get(field_id)
        if is_modification_field_empty(field_data):
            key = issue.get("key")
            if key:
                print(key)

def get_empty_modification_keys(issuesdoc, field_id):
    """Return issue keys where modification field is empty."""
    keys = []
    for issue in issuesdoc:
        field_data = issue.get("fields", {}).get(field_id)
        if is_modification_field_empty(field_data):
            key = issue.get("key")
            if key:
                keys.append(key)
    return keys

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
    global LAST_RUN_SUMMARY
    ensure_default_template()
    _log("INFO", f"Using template: {template_path}")
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

    # Separate by issue type
    stories = filter_issues_by_type(issuesdoc, STORY_TYPES)
    others = filter_issues_by_type(issuesdoc, OTHERS_TYPES)

    # Keep non-story types from being dropped if Jira introduces/uses additional names.
    known_types = STORY_TYPES.union(OTHERS_TYPES)
    uncategorized_non_story = []
    for issue in issuesdoc:
        itype = (issue.get("fields", {}).get("issuetype", {}) or {}).get("name")
        if itype and itype not in known_types:
            uncategorized_non_story.append(issue)

    if uncategorized_non_story:
        uncategorized_types = sorted({
            (issue.get("fields", {}).get("issuetype", {}) or {}).get("name")
            for issue in uncategorized_non_story
        })
        _log("INFO", f"Uncategorized non-story issue types included in others: {uncategorized_types}")
        others.extend(uncategorized_non_story)
    else:
        uncategorized_types = []

    grouped_stories = group_issues_by_heading(stories, MODIFICATION_FIELD_ID)
    grouped_others = group_issues_by_heading(others, MODIFICATION_FIELD_ID)
    story_empty_keys = get_empty_modification_keys(stories, MODIFICATION_FIELD_ID)
    other_empty_keys = get_empty_modification_keys(others, MODIFICATION_FIELD_ID)
    empty_ticket_keys = story_empty_keys + other_empty_keys
    tickets_with_modification = max(len(issuesdoc) - len(empty_ticket_keys), 0)

    # Insert for stories at <<JIRA_MODIFICATIONS>>
    inserted_stories = insert_grouped_mods_at_placeholder(
        doc,
        "<<JIRA_MODIFICATIONS>>",
        grouped_stories,
        story_empty_keys,
    )
    # Backward compatibility: if no stories placeholder found, fall back to all
    if not inserted_stories:
        fallback_empty = get_empty_modification_keys(issuesdoc, MODIFICATION_FIELD_ID)
        _log("WARN", f"Story placeholder not found; falling back to grouped-all with empty-ticket bullets: {len(fallback_empty)}")
        insert_grouped_mods_at_placeholder(doc, "<<JIRA_MODIFICATIONS>>", grouped_all, fallback_empty)

    # Insert for others at <<JIRA_MODIFICATIONS_FORDEFECTANDIMP>>
    insert_grouped_mods_at_placeholder(
        doc,
        "<<JIRA_MODIFICATIONS_FORDEFECTANDIMP>>",
        grouped_others,
        other_empty_keys,
    )

    # Save with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(OUTPUT_PATH)
    # Ensure output directory exists
    out_dir = os.path.dirname(base)
    os.makedirs(out_dir, exist_ok=True)
    output_with_timestamp = f"{base}_{timestamp}{ext}"
    doc.save(output_with_timestamp)

    cleanup_old_output_files(out_dir, os.path.basename(base), ext, OUTPUT_RETENTION_DAYS)

    summary = {
        "total_fetched": len(issuesdoc),
        "stories_total": len(stories),
        "others_total": len(others),
        "uncategorized_types": uncategorized_types,
        "tickets_with_modification": tickets_with_modification,
        "with_modification_details": sum(len(items) for items in grouped_stories.values()) + sum(len(items) for items in grouped_others.values()),
        "empty_modification_tickets": len(empty_ticket_keys),
        "story_empty_keys": story_empty_keys,
        "other_empty_keys": other_empty_keys,
        "output_file": output_with_timestamp,
    }
    LAST_RUN_SUMMARY = summary
    _log(
        "INFO",
        "Run summary: "
        f"fetched={summary['total_fetched']}, "
        f"stories={summary['stories_total']}, others={summary['others_total']}, "
        f"tickets-with-modification={summary['tickets_with_modification']}, "
        f"detailed-lines={summary['with_modification_details']}, "
        f"empty-tickets={summary['empty_modification_tickets']}",
    )

    return output_with_timestamp

def cleanup_old_output_files(directory: str, base_name: str, extension: str, retention_days: int):
    """Delete old generated docs that match the output base naming pattern."""
    if retention_days < 0:
        return

    cutoff_ts = datetime.now().timestamp() - (retention_days * 24 * 60 * 60)
    prefix = f"{base_name}_"
    deleted_count = 0

    try:
        for entry in os.listdir(directory):
            if not entry.startswith(prefix) or not entry.endswith(extension):
                continue
            path = os.path.join(directory, entry)
            if not os.path.isfile(path):
                continue
            if os.path.getmtime(path) < cutoff_ts:
                os.remove(path)
                deleted_count += 1
    except OSError as err:
        _log("WARN", f"Could not complete output cleanup: {err}")
        return

    if deleted_count:
        _log("INFO", f"Cleaned up {deleted_count} old output file(s) older than {retention_days} day(s)")

# Helper to filter issues by issuetype name
def filter_issues_by_type(issuesdoc, allowed_types: set):
    filtered = []
    for issue in issuesdoc:
        itype = (issue.get("fields", {}).get("issuetype", {}) or {}).get("name")
        if itype in allowed_types:
            filtered.append(issue)
    return filtered

# Insert grouped modifications at a given placeholder across paragraphs and tables
def insert_grouped_mods_at_placeholder(doc: Document, placeholder: str, grouped_data: dict, empty_issue_keys: list | None = None):
    inserted = False
    empty_issue_keys = empty_issue_keys or []
    empty_msg = "(No modifications found)" if not grouped_data and not empty_issue_keys else None

    def write_content(target_para):
        if empty_msg:
            target_para.add_run(empty_msg)
            return

        for heading, items in grouped_data.items():
            h = target_para
            hr = h.add_run(("\n" if h.text else "") + heading)
            hr.bold = True
            hr.underline = True
            for detail, issue_key in items:
                target_para.add_run(f"\n• {detail} ({issue_key})")

        # Keep empty modification tickets visible as blank bullets with issue key only.
        for issue_key in empty_issue_keys:
            target_para.add_run(f"\n• ({issue_key})")

    # Body paragraphs
    for para in doc.paragraphs:
        if placeholder in ("".join(run.text for run in para.runs) or para.text):
            para.text = ""
            write_content(para)
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
                            write_content(para)
                            inserted = True
                            break
                    if inserted:
                        break
                if inserted:
                    break
            if inserted:
                break
    if inserted:
        grouped_lines = sum(len(items) for items in grouped_data.values())
        _log(
            "DEBUG",
            f"[DEBUG] Inserted placeholder {placeholder}: "
            f"headings={len(grouped_data)}, detailed-lines={grouped_lines}, blank-ticket-lines={len(empty_issue_keys)}"
        )
    else:
        _log("WARN", f"Placeholder not found in template: {placeholder}")
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
    _log("INFO", f"Starting mod doc generation. JQL: {JQL_QUERY}")
    issues = fetch_issues()
    # Output keys for tickets that still need manual modification details in Jira.
    print_empty_modification_ticket_keys(issues, MODIFICATION_FIELD_ID)
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