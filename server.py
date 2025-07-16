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
</head>
<body>
  <h1>Upload MDB File</h1>
  <form method="post" enctype="multipart/form-data">
    <input type="file" name="file">
    <input type="submit" value="Upload">
  </form>
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
            finally:
                os.remove(mdb_path)
    return render_template_string(HTML_TEMPLATE, tables=tables, preview=preview)

if __name__ == '__main__':
    app.run(debug=True)
