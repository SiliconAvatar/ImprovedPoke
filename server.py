import os
import tempfile

import pyodbc
from flask import Flask, request, render_template_string, send_file, after_this_request
from openpyxl import Workbook


def export_instruments_to_excel(mdb_path: str, xlsx_path: str) -> int:
    """Export the Instruments table to an Excel workbook.

    Only rows with Type='IO' are exported. The columns Tag, FullDescription,
    EGULow, EGUHigh, RawLow, RawHigh and a set of alarm/warning columns are
    written. The instruments are grouped into DigitalInput, DigitalOutput,
    AnalogInput and AnalogOutput sheets. Returns the number of rows exported.
    """

    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={mdb_path};'
    )
    conn = pyodbc.connect(conn_str, autocommit=True)
    cursor = conn.cursor()

    # Verify the Instruments table exists
    table_names = [row.table_name for row in cursor.tables(tableType='TABLE')]
    if 'Instruments' not in table_names:
        raise ValueError('Instruments table not found')

    query = (
        "SELECT Tag, FullDescription, EGULow, EGUHigh, RawLow, RawHigh, "
        "HALM_EN, HALM_SP, HALM_DB, HALM_DLY, "
        "HWARN_EN, HWARN_SP, HWARN_DB, HWARN_DLY, "
        "LALM_EN, LALM_SP, LALM_DB, LALM_DLY, "
        "LWARN_EN, LWARN_SP, LWARN_DB, LWARN_DLY, "
        "DigitalInput, DigitalOutput, AnalogInput, AnalogOutput "
        "FROM Instruments WHERE Type='IO' AND Tag <> '' AND Tag IS NOT NULL"
    )
    cursor.execute(query)
    rows = cursor.fetchall()

    header = [
        'Tag',
        'FullDescription',
        'EGULow',
        'EGUHigh',
        'RawLow',
        'RawHigh',
        'HALM_EN',
        'HALM_SP',
        'HALM_DB',
        'HALM_DLY',
        'HWARN_EN',
        'HWARN_SP',
        'HWARN_DB',
        'HWARN_DLY',
        'LALM_EN',
        'LALM_SP',
        'LALM_DB',
        'LALM_DLY',
        'LWARN_EN',
        'LWARN_SP',
        'LWARN_DB',
        'LWARN_DLY',
    ]

    categories = {
        'DigitalInput': [],
        'DigitalOutput': [],
        'AnalogInput': [],
        'AnalogOutput': [],
    }

    for row in rows:
        if getattr(row, 'DigitalInput', False):
            categories['DigitalInput'].append(row)
        elif getattr(row, 'DigitalOutput', False):
            categories['DigitalOutput'].append(row)
        elif getattr(row, 'AnalogInput', False):
            categories['AnalogInput'].append(row)
        elif getattr(row, 'AnalogOutput', False):
            categories['AnalogOutput'].append(row)

    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    for name, rows_list in categories.items():
        ws = wb.create_sheet(name)
        ws.append(header)
        for row in rows_list:
            ws.append(row[:len(header)])

    wb.save(xlsx_path)

    conn.close()
    return sum(len(v) for v in categories.values())

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
  {% if message %}
  <p>{{ message }}</p>
  {% endif %}
  {% if excel_available %}
  <p><a href="{{ url_for('download_excel') }}">Download Instruments Excel</a></p>
  {% endif %}
</body>
</html>
"""

app = Flask(__name__)
app.config['EXCEL_PATH'] = None

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    message = None
    excel_available = False

    if request.method == 'POST':
        file = request.files.get('file')
        if file and file.filename:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.mdb') as tmp:
                file.save(tmp.name)
                mdb_path = tmp.name

            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as xlsx_tmp:
                xlsx_path = xlsx_tmp.name

            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={mdb_path};'
            )

            try:
                conn = pyodbc.connect(conn_str, autocommit=True)
                cursor = conn.cursor()
                table_names = [row.table_name for row in cursor.tables(tableType='TABLE')]

                if 'Instruments' not in table_names:
                    message = 'Uploaded file does not contain an Instruments table.'
                else:
                    count = export_instruments_to_excel(mdb_path, xlsx_path)
                    app.config['EXCEL_PATH'] = xlsx_path
                    excel_available = True
                    message = f'MDB upload successful. Found {count} valid instruments.'

                conn.close()
            except (pyodbc.Error, ValueError) as e:
                message = f'Error accessing MDB file: {e}'
                try:
                    os.remove(xlsx_path)
                except Exception:
                    pass
            finally:
                os.remove(mdb_path)

    return render_template_string(
        HTML_TEMPLATE,
        message=message,
        excel_available=excel_available,
    )


@app.route('/download_excel')
def download_excel():
    """Send the exported instruments Excel file to the client."""
    xlsx_path = app.config.get('EXCEL_PATH')
    if not xlsx_path or not os.path.exists(xlsx_path):
        return "No Excel file available", 404

    @after_this_request
    def remove_file(response):
        try:
            os.remove(xlsx_path)
        except Exception:
            pass
        app.config['EXCEL_PATH'] = None
        return response

    return send_file(xlsx_path, as_attachment=True, download_name='instruments.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
