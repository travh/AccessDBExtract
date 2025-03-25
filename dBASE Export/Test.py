import pyodbc
import pandas as pd

def list_columns_and_types(db_path, table_name):
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + db_path + ';'
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Read the table into a DataFrame
    query = f"SELECT * FROM {table_name}"
    df = pd.read_sql(query, conn)
    
    # Convert "text" columns to string data type
    for column in df.columns:
        if df[column].dtype == 'object':
            df[column] = df[column].astype('string')
    
    # List the columns, their Access data types, and df.dtype
    columns_info = []
    for column in df.columns:
        if df[column].dtype == 'int64':
            access_data_type = 'Long Integer'
        elif df[column].dtype == 'float64':
            access_data_type = 'Double'
        elif df[column].dtype == 'bool':
            access_data_type = 'Yes/No'
        elif df[column].dtype == 'datetime64[ns]':
            access_data_type = 'Date/Time'
        elif df[column].dtype == 'timedelta64[ns]':
            access_data_type = 'Date/Time'
        else:
            access_data_type = 'Text'
        
        columns_info.append((column, access_data_type, df[column].dtype))

    # Print the columns, their Access data types, and df.dtype
    for column, access_data_type, dtype in columns_info:
        print(f'Column: {column}, Access Data Type: {access_data_type}, df.dtype: {dtype}')

    cursor.close()
    conn.close()

# Set the path to the Access database and the table name
db_path = r"C:\AccessTest\DB1\Database.accdb"
table_name = "Customers"

# Run the process to list columns and their types
list_columns_and_types(db_path, table_name)