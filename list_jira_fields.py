import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

JIRA_BASE_URL = os.getenv("JIRA_BASE_URL")
JIRA_EMAIL = os.getenv("JIRA_EMAIL")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")

auth = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)
headers = {
    "Accept": "application/json"
}

def list_custom_fields():
    url = f"{JIRA_BASE_URL}/rest/api/3/field"
    response = requests.get(url, headers=headers, auth=auth)
    response.raise_for_status()
    fields = response.json()
    for field in fields:
        schema = field.get('schema', {})
        print(f"Name: {field['name']} | ID: {field['id']} | Type: {schema.get('type', 'N/A')}")

if __name__ == "__main__":
    list_custom_fields()
