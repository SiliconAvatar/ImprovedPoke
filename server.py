import os
import tempfile

import pyodbc
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
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={mdb_path};'
            )
            try:
                conn = pyodbc.connect(conn_str, autocommit=True)
                cursor = conn.cursor()
                tables = [row.table_name for row in cursor.tables(tableType='TABLE')]
                if tables:
                    first = tables[0]
                    cursor.execute(f'SELECT * FROM [{first}]')
                    rows = cursor.fetchmany(5)
                    columns = [col[0] for col in cursor.description]
                    lines = [', '.join(columns)]
                    for r in rows:
                        lines.append(', '.join(str(item) for item in r))
                    preview = {
                        'table': first,
                        'data': '\n'.join(lines)
                    }
                conn.close()
            except pyodbc.Error as e:
                preview = {
                    'table': 'Error',
                    'data': f'Error accessing MDB file: {e}'
                }
            finally:
                os.remove(mdb_path)
    return render_template_string(HTML_TEMPLATE, tables=tables, preview=preview)

if __name__ == '__main__':
    app.run(debug=True)
