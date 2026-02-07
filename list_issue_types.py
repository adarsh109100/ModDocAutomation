import os
import json
import sys
import requests
from requests.auth import HTTPBasicAuth

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "jira_config.json")


def read_config():
    if not os.path.exists(CONFIG_PATH):
        print(f"Config file not found: {CONFIG_PATH}")
        return None
    with open(CONFIG_PATH, 'r') as f:
        return json.load(f)


def list_all_issue_types(base_url: str, email: str, token: str):
    """Calls Jira API to list all global issue types and prints them."""
    url = f"{base_url.rstrip('/')}/rest/api/3/issuetype"
    auth = HTTPBasicAuth(email, token)
    headers = {"Accept": "application/json"}
    try:
        resp = requests.get(url, headers=headers, auth=auth, timeout=15)
        print(f"[DEBUG] GET {url} -> {resp.status_code}")
        if resp.status_code != 200:
            print(resp.text)
            resp.raise_for_status()
        data = resp.json()
        if isinstance(data, list):
            print("\nIssue Types:")
            for it in data:
                print(f"- Name: {it.get('name')} | ID: {it.get('id')} | Subtask: {it.get('subtask')}")
        else:
            print("Unexpected response format:", data)
    except requests.exceptions.RequestException as e:
        print("Network error:", e)


def list_issue_types_from_search(base_url: str, email: str, token: str, jql: str):
    """Optionally, search with JQL and print issuetype names appearing in the result set."""
    url = f"{base_url.rstrip('/')}/rest/api/3/search/jql"
    auth = HTTPBasicAuth(email, token)
    headers = {"Accept": "application/json"}
    params = {"jql": jql, "fields": "issuetype", "maxResults": 100}
    try:
        resp = requests.get(url, headers=headers, auth=auth, params=params, timeout=15)
        print(f"[DEBUG] GET {url} -> {resp.status_code}")
        if resp.status_code != 200:
            print(resp.text)
            resp.raise_for_status()
        issues = resp.json().get("issues", [])
        types = {}
        for issue in issues:
            it = (issue.get("fields", {}).get("issuetype", {}) or {}).get("name")
            if it:
                types[it] = types.get(it, 0) + 1
        print("\nIssue types in search results:")
        for name, count in sorted(types.items()):
            print(f"- {name}: {count}")
    except requests.exceptions.RequestException as e:
        print("Network error:", e)


if __name__ == "__main__":
    cfg = read_config()
    if not cfg:
        sys.exit(1)
    base = cfg.get("JIRA_BASE_URL", "")
    email = cfg.get("JIRA_EMAIL", "")
    token = cfg.get("JIRA_API_TOKEN", "")
    if not base or not email or not token:
        print("Please populate JIRA_BASE_URL, JIRA_EMAIL, JIRA_API_TOKEN in jira_config.json")
        sys.exit(1)

    print("\nListing all global issue types:\n")
    list_all_issue_types(base, email, token)

    # Optional: pass a JQL via env or default
    jql = os.getenv("JQL_QUERY", "project = OLK ORDER BY created DESC")
    print("\nListing issue types present in a sample search:\n")
    list_issue_types_from_search(base, email, token, jql)
