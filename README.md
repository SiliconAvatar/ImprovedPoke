# MDB Web Server

This repository contains a small Flask application that allows uploading a
Microsoft Access `.mdb` file. After uploading, the server lists the tables
found in the database and shows a preview of the first table using
`mdbtools`.

## Requirements

* Python 3.11+
* `Flask` Python package
* `mdbtools` command line utilities (`mdb-tables`, `mdb-export`)

You can install the Python dependencies with:

```bash
pip install flask
```

On Debian/Ubuntu systems `mdbtools` can be installed via `apt`:

```bash
sudo apt-get install mdbtools
```

## Running

Start the server with:

```bash
python server.py
```

Then open your browser to [http://localhost:5000](http://localhost:5000)
and upload an `.mdb` file to view its tables.
