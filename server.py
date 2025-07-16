import os
import subprocess
import tempfile
from flask import Flask, request, render_template_string

HTML_TEMPLATE = """
<!doctype html>
<html lang='en'>
<head>
  <meta charset='utf-8'>
  <title>MDB Reader</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; }
    #file-name { margin-left: 10px; }
  </style>
</head>
<body>
  <h1>Upload MDB File</h1>
  <form method="post" enctype="multipart/form-data">
    <input id="mdb-file" type="file" name="file" accept=".mdb" style="display:none" />
    <button type="button" onclick="document.getElementById('mdb-file').click()">Select File</button>
    <span id="file-name"></span>
    <button type="submit">Upload</button>
  </form>
  <script>
    const input = document.getElementById('mdb-file');
    const fileName = document.getElementById('file-name');
    input.addEventListener('change', () => {
      const file = input.files[0];
      fileName.textContent = file ? file.name : '';
    });
  </script>
  {% if tables %}
  <h2>Tables</h2>
  <ul>
  {% for t in tables %}
    <li>{{ t }}</li>
  {% endfor %}
  </ul>
  {% endif %}
  {% if preview %}
  <h2>Preview of {{ preview.table }}</h2>
  <pre>{{ preview.data }}</pre>
  {% endif %}
</body>
</html>
"""

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    tables = []
    preview = None
    if request.method == 'POST':
        file = request.files.get('file')
        if file and file.filename:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.mdb') as tmp:
                file.save(tmp.name)
                mdb_path = tmp.name
            try:
                output = subprocess.check_output(['mdb-tables', '-1', mdb_path], text=True)
                tables = [t for t in output.splitlines() if t]
                if tables:
                    first = tables[0]
                    data = subprocess.check_output(['mdb-export', mdb_path, first], text=True)
                    preview = {
                        'table': first,
                        'data': '\n'.join(data.splitlines()[:5])
                    }
            except FileNotFoundError:
                preview = {
                    'table': 'Error',
                    'data': 'mdbtools utilities not found. Please install mdbtools and ensure it is in your PATH.'
                }
            finally:
                os.remove(mdb_path)
    return render_template_string(HTML_TEMPLATE, tables=tables, preview=preview)

if __name__ == '__main__':
    app.run(debug=True)
