import csv
import os
import sys

import pyodbc


def main():
    if len(sys.argv) < 2:
        print('Usage: python export_instruments.py <path_to_mdb>')
        return

    mdb_path = sys.argv[1]
    if not os.path.exists(mdb_path):
        print(f'MDB file not found: {mdb_path}')
        return

    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={mdb_path};'
    )
    try:
        conn = pyodbc.connect(conn_str, autocommit=True)
        cursor = conn.cursor()
        query = (
            "SELECT Tag, FullDescription, EGULow, EGUHigh, RawLow, RawHigh "
            "FROM Instruments WHERE Type='IO'"
        )
        cursor.execute(query)
        rows = cursor.fetchall()
    except pyodbc.Error as e:
        print(f'Error accessing MDB file: {e}')
        return
    finally:
        try:
            conn.close()
        except Exception:
            pass

    output_path = input('Enter path to save CSV file: ').strip()
    if not output_path:
        print('No output path provided')
        return

    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Tag', 'FullDescription', 'EGULow', 'EGUHigh', 'RawLow', 'RawHigh'])
        for row in rows:
            writer.writerow(row)
    print(f'Exported {len(rows)} rows to {output_path}')


if __name__ == '__main__':
    main()
