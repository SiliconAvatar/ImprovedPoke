# MDB Web Server

This repository contains a small Flask application that allows uploading a
Microsoft Access `.mdb` file. After uploading, the server exports the
`Instruments` table to CSV and reports how many valid instruments were found
using the Microsoft Access ODBC driver via `pyodbc`.

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
and upload an `.mdb` file. After uploading, the server exports the
`Instruments` table to a CSV file and presents a link to download the file,
reporting how many instruments were found.

## Exporting Instrument Data

The `export_instruments.py` script allows exporting selected columns from the
`Instruments` table of an MDB file. It filters rows where the `Type` column is
`IO` **and where the `Tag` column is not blank**. The script writes the columns
`Tag`, `FullDescription`, `EGULow`, `EGUHigh`, `RawLow`, `RawHigh` and several
alarm/warning fields to a CSV file. The additional columns are `HALM_EN`,
`HALM_SP`, `HALM_DB`, `HALM_DLY`, `HWARN_EN`, `HWARN_SP`, `HWARN_DB`,
`HWARN_DLY`, `LALM_EN`, `LALM_SP`, `LALM_DB`, `LALM_DLY`, `LWARN_EN`,
`LWARN_SP`, `LWARN_DB`, and `LWARN_DLY`.

Usage:

```bash
python export_instruments.py path/to/database.mdb
```

After running, the script prompts for the destination path of the CSV file.
