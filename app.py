from flask import Flask, render_template_string, request, send_file
import os
from JIraOLK import generate_mod_doc

app = Flask(__name__)

FORM_HTML = '''
<!doctype html>
<html lang="en">
  <head>
    <title>JIRA Mod Doc Generator</title>
  </head>
  <body>
    <h2>JIRA Mod Doc Generator</h2>
    <form method="post">
      <label for="jira_query">JIRA Query:</label><br>
      <input type="text" id="jira_query" name="jira_query" size="80" required><br><br>
      <button type="submit">Create Mod Doc</button>
    </form>
    {% if file_url %}
      <p><a href="{{ file_url }}">Download Mod Doc</a></p>
    {% endif %}
  </body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def index():
    file_url = None
    if request.method == 'POST':
        jira_query = request.form['jira_query']
        output_path = generate_mod_doc(jira_query)
        filename = os.path.basename(output_path)
        file_url = f'/download/{filename}'
        # Store the path for download
        app.config['LAST_OUTPUT_PATH'] = output_path
    return render_template_string(FORM_HTML, file_url=file_url)

@app.route('/download/<filename>')
def download(filename):
    output_path = app.config.get('LAST_OUTPUT_PATH')
    if output_path and os.path.basename(output_path) == filename:
        return send_file(output_path, as_attachment=True)
    return "File not found", 404

if __name__ == '__main__':
    app.run(debug=True)

