import os
import csv
import tempfile

import pyodbc
from flask import Flask, request, render_template_string, send_file, after_this_request


def export_instruments_to_csv(mdb_path: str, csv_path: str) -> int:
    """Export the Instruments table to a CSV file.

    Only rows with Type='IO' are exported and the columns Tag,
    FullDescription, EGULow, EGUHigh, RawLow and RawHigh are written.
    Returns the number of rows exported.
    """
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={mdb_path};'
    )
    conn = pyodbc.connect(conn_str, autocommit=True)
    cursor = conn.cursor()
    query = (
        "SELECT Tag, FullDescription, EGULow, EGUHigh, RawLow, RawHigh "
        "FROM Instruments WHERE Type='IO' AND Tag <> '' AND Tag IS NOT NULL"
    )
    cursor.execute(query)
    rows = cursor.fetchall()
    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow([
            'Tag',
            'FullDescription',
            'EGULow',
            'EGUHigh',
            'RawLow',
            'RawHigh',
        ])
        for row in rows:
            writer.writerow(row)
    conn.close()
    return len(rows)

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
  {% if csv_available %}
  <p><a href="{{ url_for('download_csv') }}">Download Instruments CSV</a></p>
  {% endif %}
</body>
</html>
"""

app = Flask(__name__)
app.config['CSV_PATH'] = None

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    tables = []
    preview = None
    csv_available = False
    if request.method == 'POST':
        file = request.files.get('file')
        if file and file.filename:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.mdb') as tmp:
                file.save(tmp.name)
                mdb_path = tmp.name
            with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as csv_tmp:
                csv_path = csv_tmp.name
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
                export_instruments_to_csv(mdb_path, csv_path)
                app.config['CSV_PATH'] = csv_path
                csv_available = True
            except pyodbc.Error as e:
                preview = {
                    'table': 'Error',
                    'data': f'Error accessing MDB file: {e}'
                }
                try:
                    os.remove(csv_path)
                except Exception:
                    pass
            finally:
                os.remove(mdb_path)
    return render_template_string(
        HTML_TEMPLATE,
        tables=tables,
        preview=preview,
        csv_available=csv_available,
    )


@app.route('/download_csv')
def download_csv():
    """Send the exported instruments CSV to the client."""
    csv_path = app.config.get('CSV_PATH')
    if not csv_path or not os.path.exists(csv_path):
        return "No CSV file available", 404

    @after_this_request
    def remove_file(response):
        try:
            os.remove(csv_path)
        except Exception:
            pass
        app.config['CSV_PATH'] = None
        return response

    return send_file(csv_path, as_attachment=True, download_name='instruments.csv')

if __name__ == '__main__':
    app.run(debug=True)
