import os
import pyodbc

# Function to get all table names from MS Access database
def get_table_names(db_path):
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    table_names = [row.table_name for row in cursor.tables(tableType='TABLE')]
    conn.close()
    return table_names

# Function to list tables in each database
def list_tables_in_databases(db_folder):
    db_tables = {}
    for db_file in os.listdir(db_folder):
        if db_file.endswith('.accdb'):
            db_path = os.path.join(db_folder, db_file)
            table_names = get_table_names(db_path)
            db_tables[db_file] = table_names
    return db_tables

# Example usage
db_folder = 'C:\AccessTest\DB1'
db_tables = list_tables_in_databases(db_folder)

for db_file, tables in db_tables.items():
    print(f"Database: {db_file}")
    for table in tables:
        print(f"  Table: {table}")

print("Listing of tables in each database is complete.")