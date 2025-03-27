import os
import pyodbc
import pandas as pd
import warnings
from sqlalchemy import create_engine

# Suppress specific UserWarning
warnings.filterwarnings("ignore", category=UserWarning, message="pandas only supports SQLAlchemy connectable")

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

# Function to write data to SQL Server database
def write_data_to_sql(df, table_name, sql_conn_str):
    # Create SQLAlchemy engine
    engine = create_engine(sql_conn_str)
    
    # Write data to SQL Server (create table if it doesn't exist)
    df.to_sql(table_name, engine, if_exists='replace', index=False)

# Function to traverse directories and find all Access databases
def find_access_databases(folder):
    access_databases = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('.accdb'):
                access_databases.append(os.path.join(root, file))
    return access_databases

# Example usage
db_folder = 'C:\\AccessTest'
sql_conn_str = 'mssql+pyodbc://your_username:your_password@your_server_name/your_database_name?driver=ODBC+Driver+17+for+SQL+Server'

access_databases = find_access_databases(db_folder)

total_databases = len(access_databases)
for i, db_path in enumerate(access_databases):
    print(f"Extracting from database {i + 1} of {total_databases}")
    db_folder_name = os.path.basename(os.path.dirname(db_path))
    table_names = get_table_names(db_path)
    for table_name in table_names:
        try:
            data = extract_data_from_access(db_path, table_name)
            # Add the DatabaseID column with the folder name as its value
            data.insert(0, 'DatabaseID', db_folder_name)
            write_data_to_sql(data, table_name, sql_conn_str)
        except Exception as e:
            print(f"Error processing table {table_name} in database {db_path}: {e}")

print("Data consolidation to SQL Server is complete.")