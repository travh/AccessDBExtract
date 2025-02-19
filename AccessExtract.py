import os
import pyodbc
import pandas as pd

# Function to get all table names from MS Access database
def get_table_names(db_path):
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    table_names = [row.table_name for row in cursor.tables(tableType='TABLE')]
    conn.close()
    return table_names

# Function to extract data from a specific table in MS Access database
def extract_data_from_access(db_path, table_name):
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path
    conn = pyodbc.connect(conn_str)
    query = f'SELECT * FROM [{table_name}]'
    df = pd.read_sql(query, conn)
    conn.close()
    return df

# Function to consolidate data into CSV files
def consolidate_data_to_csv(df, table_name, output_folder, db_file):
    # Add a column for the database name
    df.insert(0, 'DatabaseName', db_file)
    
    # Replace spaces in table names with underscores for the CSV file name
    sanitized_table_name = table_name.replace(" ", "_")
    output_path = os.path.join(output_folder, f'{sanitized_table_name}.csv')
    
    if os.path.exists(output_path):
        df.to_csv(output_path, mode='a', header=False, index=False)
    else:
        df.to_csv(output_path, index=False)
        
# Example usage
db_folder = 'C:\AccessTest'
output_folder = 'C:\AccessTest\Output'

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

for db_file in os.listdir(db_folder):
    if db_file.endswith('.accdb'):
        db_path = os.path.join(db_folder, db_file)
        table_names = get_table_names(db_path)
        for table_name in table_names:
            try:
                data = extract_data_from_access(db_path, table_name)
                consolidate_data_to_csv(data, table_name, output_folder, db_file)
            except Exception as e:
                print(f"Error processing table {table_name} in database {db_file}: {e}")

print("Data consolidation to CSV files is complete.")