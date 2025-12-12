import os
import platform


def read_table(table_name):
    if platform.system() == 'Windows':
        return read_table_windows(table_name)
    elif platform.system() == 'Darwin':
        return read_table_macos(table_name)
    else:
        raise Exception("Unsupported OS")

def read_table_windows(table_name):
    import pyodbc
    db_path = os.getenv("DATABASE_PATH")
    conn = pyodbc.connect(f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path}")

    cursor = conn.cursor()
    query = f"SELECT * FROM {table_name}"
    cursor.execute(query)
    columns = [column[0] for column in cursor.description]
    rows = cursor.fetchall()

    table_dict = {column: [] for column in columns}
    for row in rows:
        for col, value in zip(columns, row):
            table_dict[col].append(value)
    
    conn.close()
    return table_dict


def read_table_macos(table_name):
    from access_parser import AccessParser
    db_path = os.getenv("DATABASE_PATH")
    db = AccessParser(db_path)
    return db.parse_table(table_name)

