import pyodbc
import sqlite3


def mdb_to_sqlite(mdb_path, sqlite_path):
    # Connect to the .mdb file using ODBC
    conn_mdb = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + mdb_path)

    cursor_mdb = conn_mdb.cursor()

    # Create a new SQLite database and establish a connection
    conn_sqlite = sqlite3.connect(sqlite_path)
    cursor_sqlite = conn_sqlite.cursor()

    # Retrieve table names from the MDB file
    tables = [row.table_name for row in cursor_mdb.tables(tableType='TABLE')]

    for table in tables:
        print(f"Processing table: {table}")

        # Attempt to retrieve column details
        try:
            columns = [x.column_name for x in cursor_mdb.columns(table=table)]
            columns_declaration = ', '.join(
                [f'"{col}" TEXT' for col in columns])  # Ensuring column names are handled as strings
        except UnicodeDecodeError as e:
            print(f"Error decoding column names for table {table}: {e}")
            continue  # Skip this table or handle differently if necessary

        # Create table in SQLite
        cursor_sqlite.execute(f"CREATE TABLE IF NOT EXISTS \"{table}\" ({columns_declaration})")

        # Select data from Access table
        try:
            cursor_mdb.execute(f"SELECT * FROM [{table}]")
            rows = cursor_mdb.fetchall()

            # Insert data into SQLite table
            placeholders = ', '.join(['?' for _ in columns])
            cursor_sqlite.executemany(f"INSERT INTO \"{table}\" VALUES ({placeholders})", rows)
        except Exception as e:
            print(f"Failed to process table {table}: {e}")
            continue  # Handle the error appropriately

        # Commit changes to SQLite database
        conn_sqlite.commit()

    # Close all connections
    cursor_mdb.close()
    conn_mdb.close()
    cursor_sqlite.close()
    conn_sqlite.close()

    print("Conversion completed.")

# Path to your .mdb and .sqlite files
mdb_path = 'C:/Users/yannm/Dropbox/04 PhD/06 Projects/Oases/03 Code/01_build/input/hwsd/HWSD2.mdb'
sqlite_path = 'C:/Users/yannm/Dropbox/04 PhD/06 Projects/Oases/03 Code/01_build/input/hwsd/sqlite/hwsd2.sqlite'

# Call the function to perform the conversion
mdb_to_sqlite(mdb_path, sqlite_path)