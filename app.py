from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os
import json
import requests
from requests.auth import HTTPBasicAuth
from JIraOLK import generate_mod_doc, set_jira_config

app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = "change-me"
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "jira_config.json")


def read_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r') as f:
            return json.load(f)
    return {
        "JIRA_BASE_URL": "",
        "JIRA_EMAIL": "",
        "JIRA_API_TOKEN": "",
    }


def write_config(cfg):
    with open(CONFIG_PATH, 'w') as f:
        json.dump(cfg, f)


@app.route('/', methods=['GET'])
def home():
    cfg = read_config()
    return render_template('index.html', config=cfg)


@app.route('/save-config', methods=['POST'])
def save_config():
    cfg = {
        "JIRA_BASE_URL": request.form.get('JIRA_BASE_URL', '').rstrip('/'),
        "JIRA_EMAIL": request.form.get('JIRA_EMAIL', ''),
        "JIRA_API_TOKEN": request.form.get('JIRA_API_TOKEN', ''),
    }
    write_config(cfg)
    set_jira_config(cfg)
    flash('Configuration saved.')
    return redirect(url_for('home'))


@app.route('/test-connection', methods=['POST'])
def test_connection():
    cfg = read_config()
    base = cfg.get('JIRA_BASE_URL', '').rstrip('/')
    email = cfg.get('JIRA_EMAIL', '')
    token = cfg.get('JIRA_API_TOKEN', '')
    if not base or not email or not token:
        flash('Please fill JIRA Base URL, Email, and API Token before testing.')
        return redirect(url_for('home'))
    try:
        url = f"{base}/rest/api/3/myself"
        headers = {"Accept": "application/json"}
        auth = HTTPBasicAuth(email, token)
        resp = requests.get(url, headers=headers, auth=auth, timeout=10)
        if resp.status_code == 200:
            me = resp.json()
            flash(f"Connection successful. Authenticated as: {me.get('displayName', email)}")
        else:
            flash(f"Connection failed. Status {resp.status_code}: {resp.text}")
    except requests.exceptions.RequestException as e:
        flash(f"Connection error: {e}")
    return redirect(url_for('home'))


@app.route('/generate-and-download', methods=['POST'])
def generate_and_download():
    cfg = read_config()
    set_jira_config(cfg)

    jira_query = request.form.get('jira_query', '').strip()
    meta = {
        "version": request.form.get('version', '').strip(),
        "release": request.form.get('release', '').strip(),
        "revision": request.form.get('revision', '').strip(),
        "build": request.form.get('build', '').strip(),
    }

    if not jira_query:
        flash('Please enter a JQL query.')
        return redirect(url_for('home'))

    # generate_mod_doc should accept meta for placeholder replacement
    output_path = generate_mod_doc(jira_query, meta)

    if output_path and os.path.exists(output_path):
        return send_file(output_path, as_attachment=True)

    flash('Failed to generate Mod Doc')
    return redirect(url_for('home'))


if __name__ == '__main__':
    app.run(debug=True)
