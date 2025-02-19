import pyodbc

# Function to get the schema of a specific MS Access database
def get_database_schema(db_path):
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    schema = {}
    for row in cursor.tables(tableType='TABLE'):
        table_name = row.table_name if hasattr(row, 'table_name') else row
        cursor.execute(f'SELECT * FROM [{table_name}] WHERE 1=0')
        columns = [column for column in cursor.description]
        schema[table_name] = columns
    conn.close()
    return schema

# Example usage
db_path = 'C:\AccessTest\Database1.accdb'
schema = get_database_schema(db_path)

for table_name, columns in schema.items():
    print(f"Table: {table_name}")
    for column in columns:
        print(f"  Column: {column}")

print("Schema extraction is complete.")