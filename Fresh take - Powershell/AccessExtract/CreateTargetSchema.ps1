# Define the path to the MS Access database
$accessDatabasePath = "C:\AccessTest\DB1\Database.accdb"  # Replace this with your database path

# Define the connection string for SQL Server
$sqlServerInstance = "DESKTOP-VKMSDNG"  # Replace with your SQL Server instance name
$sqlDatabaseName = "AccessStaging"  # Replace with your target SQL Server database name
$sqlConnectionString = "Server=$sqlServerInstance;Database=$sqlDatabaseName;Integrated Security=True;"  # Adjust if using SQL authentication

# Create a new OleDbConnection object for MS Access
$accessConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$accessDatabasePath;"
$accessConnection = New-Object System.Data.OleDb.OleDbConnection($accessConnectionString)

# Create a new SqlConnection object for SQL Server
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection($sqlConnectionString)

# Open the MS Access connection
$accessConnection.Open()

# Open the SQL Server connection
$sqlConnection.Open()

# Get the schema information for the tables in the MS Access database
$accessTables = $accessConnection.GetSchema("Tables")

# Function to map MS Access data types to SQL Server data types
function Get-SqlServerDataType {
    param ($accessDataType)

    switch ($accessDataType) {
        "TEXT" { return "VARCHAR(MAX)" }
        "MEMO" { return "TEXT" }
        "INTEGER" { return "INT" }
        "LONG" { return "BIGINT" }
        "DOUBLE" { return "FLOAT" }
        "CURRENCY" { return "MONEY" }
        "DATE" { return "DATETIME" }
        "YESNO" { return "BIT" }
        "COUNTER" { return "INT" }
        default { return "VARCHAR(MAX)" }  # Fallback for unknown types
    }
}

# Function to delimit table and column names with square brackets
function Delimit-SqlName {
    param ($name)
    return "[$name]"  # Simply wrap all names in square brackets
}

# Loop through each MS Access table and check if it exists in SQL Server
foreach ($accessTable in $accessTables.Rows) {
    $tableName = $accessTable["TABLE_NAME"]
    
    # Skip tables whose names start with "MSys"
    if ($tableName -like "MSys*") {
        continue
    }

    $delimitedTableName = Delimit-SqlName $tableName
    
    # Check if the table exists in SQL Server
    $checkTableQuery = "SELECT CASE WHEN EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '$tableName') THEN 1 ELSE 0 END"
    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand($checkTableQuery, $sqlConnection)
    
    # Execute the query
    $tableExists = $sqlCommand.ExecuteScalar()

    if ($tableExists -eq 1) {
        # Table already exists, so skip creation
        Write-Output "Table $tableName already exists in SQL Server. Skipping creation."
        continue
    }

    # If table doesn't exist, create it
    Write-Output "Creating table $tableName in SQL Server."

    # Get the columns and data types from the MS Access table without restrictions
    $columns = $accessConnection.GetSchema("Columns")

    # Filter columns for the current table
    $columnsForTable = $columns | Where-Object { $_["TABLE_NAME"] -eq $tableName }

    # Start building the CREATE TABLE SQL statement
    $createTableSql = "CREATE TABLE $delimitedTableName ("

    $columnDefinitions = @()

    foreach ($column in $columnsForTable) {
        $columnName = $column["COLUMN_NAME"]
        $delimitedColumnName = Delimit-SqlName $columnName
        $accessDataType = $column["DATA_TYPE"]

        # Map the MS Access data type to SQL Server data type
        $sqlServerDataType = Get-SqlServerDataType $accessDataType

        # Build the column definition
        $columnDefinitions += "$delimitedColumnName $sqlServerDataType"
    }

    # Add TempInvestigationID column
    $columnDefinitions += "[TempInvestigationID] NVARCHAR(500)"

    # Combine all column definitions and add closing parenthesis
    $createTableSql += ($columnDefinitions -join ", ") + ")"

    try {
        # Attempt to execute the CREATE TABLE SQL command
        $createTableCommand = New-Object System.Data.SqlClient.SqlCommand($createTableSql, $sqlConnection)
        $createTableCommand.ExecuteNonQuery() | Out-Null

        Write-Output "Table $tableName created in SQL Server."

    } catch {
        # If there is an error, output the SQL statement and the error message
        Write-Output "Failed to create table $tableName. SQL: $createTableSql"
        Write-Output "Error: $($_.Exception.Message)"
    }
}

# Close the connections
$accessConnection.Close()
$sqlConnection.Close()
