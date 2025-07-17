# MDB Web Server

This repository contains a small Flask application that allows uploading a
Microsoft Access `.mdb` file. After uploading, the server lists the tables
found in the database and shows a preview of the first table using the
Microsoft Access ODBC driver via `pyodbc`.

## Requirements

* Python 3.11+
* `Flask` Python package
* `pyodbc` package
* Microsoft Access ODBC driver (installed with Microsoft Office/Access)

You can install the Python dependencies with:

```bash
pip install flask pyodbc
```

This application requires the Microsoft Access ODBC driver. On Windows this
driver is installed when Microsoft Access or the Access Database Engine is
present. Ensure the driver name `Microsoft Access Driver (*.mdb, *.accdb)` is
available on your system.

## Running

Start the server with:

```bash
python server.py
```

Then open your browser to [http://localhost:5000](http://localhost:5000)
and upload an `.mdb` file to view its tables. After uploading the server also
exports the `Instruments` table to a CSV file and presents a link to download
the file.

## Exporting Instrument Data

The `export_instruments.py` script allows exporting selected columns from the
`Instruments` table of an MDB file. It filters rows where the `Type` column is
`IO` **and where the `Tag` column is not blank**. The script writes the columns
`Tag`, `FullDescription`, `EGULow`, `EGUHigh`, `RawLow`, and `RawHigh` to a CSV
file.

Usage:

```bash
python export_instruments.py path/to/database.mdb
```

After running, the script prompts for the destination path of the CSV file.
