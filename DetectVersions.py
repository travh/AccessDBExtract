import os
import pyodbc

# Function to get all table names and their columns from MS Access database
def get_table_schemas(db_path):
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    table_schemas = {}
    for row in cursor.tables(tableType='TABLE'):
        table_name = row.table_name if hasattr(row, 'table_name') else row
        if table_name not in table_schemas:
            try:
                cursor.execute(f'SELECT * FROM [{table_name}] WHERE 1=0')
                columns = [column for column in cursor.description]
                table_schemas[table_name] = columns
            except pyodbc.ProgrammingError:
                continue
    conn.close()
    return table_schemas

# Function to assign versions to databases based on schema differences
def assign_database_versions(db_folder):
    db_versions = {}
    version_schemas = []
    version_counter = 1

    for db_file in os.listdir(db_folder):
        if db_file.endswith('.accdb'):
            db_path = os.path.join(db_folder, db_file)
            table_schemas = get_table_schemas(db_path)

            # Check if the schema matches any existing version
            schema_matched = False
            for version, schema in enumerate(version_schemas, start=1):
                if table_schemas == schema:
                    db_versions[db_file] = version
                    schema_matched = True
                    break

            # If no match, assign a new version
            if not schema_matched:
                db_versions[db_file] = version_counter
                version_schemas.append(table_schemas)
                version_counter += 1

    return db_versions

# Example usage
db_folder = 'C:\AccessTest'
db_versions = assign_database_versions(db_folder)

for db_file, version in db_versions.items():
    print(f"Database: {db_file} - Version: {version}")

print("Database version assignment is complete.")