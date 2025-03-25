import pyodbc
import pandas as pd
import dbf

def export_table_to_dbf(db_path, table_name, export_file):
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
    
    # Define the fields for the dBASE table and identify the column causing the issue
    fields = []
    for column in df.columns:
        try:
            if df[column].dtype == 'int64':
                fields.append((column, dbf.Integer))  # Use dbf.Integer for int64
            elif df[column].dtype == 'float64':
                fields.append((column, dbf.Float))  # Use dbf.Float for float64
            elif df[column].dtype == 'bool':
                fields.append((column, dbf.Logical))
            elif df[column].dtype == 'datetime64[ns]':
                fields.append((column, dbf.DateTime))
            elif df[column].dtype == 'timedelta64[ns]':
                fields.append((column, dbf.Time))
            else:
                fields.append((column, dbf.Char))
        except Exception as e:
            print(f'Error with column {column}: {e}')

    # Print the defined fields
    print(fields)
    
    # Create a new dBASE table
    table = dbf.Table(export_file, fields)
    table.open(dbf.READ_WRITE)
    
    # Insert data into the dBASE table
    for index, row in df.iterrows():
        try:
            # Convert each value in the row to a string if necessary
            row = tuple(str(value) if pd.notnull(value) else '' for value in row)  # Handle NaN values as empty strings
            table.append(row)
        except Exception as e:
            print(f'Error appending row {index}: {e}')
            print(f'Row data: {row}')
    
    table.close()
    print(f'Exported {table_name} to {export_file}')

    cursor.close()
    conn.close()

# Set the path to the Access database, the table name, and the export file path
db_path = r"C:\AccessTest\DB1\Database.accdb"
table_name = "Customers"
export_file = r"C:\AccessTest\Output\dBASE\exported_table.dbf"

# Run the export process
export_table_to_dbf(db_path, table_name, export_file)