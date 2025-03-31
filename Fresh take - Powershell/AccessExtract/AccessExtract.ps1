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

# Loop through each MS Access table and check if it exists in SQL Server
foreach ($accessTable in $accessTables.Rows) {
    $tableName = $accessTable["TABLE_NAME"]
    
    # Check if the table exists in SQL Server
    $checkTableQuery = "IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '$tableName') SELECT 1 ELSE SELECT 0"
    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand($checkTableQuery, $sqlConnection)
    
    # Execute the query
    $tableExists = $sqlCommand.ExecuteScalar()

    # Output the table name and whether it exists in SQL Server
    $exists = if ($tableExists -eq 1) { $true } else { $false }
    Write-Output "${tableName}: $exists"
}

# Close the connections
$accessConnection.Close()
$sqlConnection.Close()