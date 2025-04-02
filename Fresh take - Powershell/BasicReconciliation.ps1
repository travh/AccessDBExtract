# Define the root folder containing MS Access databases
$accessRootFolder = "C:\AccessTest"  # Change this to your root folder
$accessDatabaseName = "database.accdb"  # Name of the Access database to match

# Define the connection string for SQL Server
$sqlServerInstance = "DESKTOP-VKMSDNG"
$sqlDatabaseName = "AccessStaging"
$sqlConnectionString = "Server=$sqlServerInstance;Database=$sqlDatabaseName;Integrated Security=True;"

# Create a new SqlConnection object for SQL Server
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection($sqlConnectionString)
$sqlConnection.Open()

# Get all Access database files in the root folder and subfolders matching the specific database name
$accessDatabaseFiles = Get-ChildItem -Path $accessRootFolder -Recurse | Where-Object { $_.Name -eq $accessDatabaseName }

foreach ($accessDatabase in $accessDatabaseFiles) {
    Write-Output "Reconciling database: $($accessDatabase.FullName)"
    
    # Define the connection string for the current Access database
    $accessConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$($accessDatabase.FullName);"
    $accessConnection = New-Object System.Data.OleDb.OleDbConnection($accessConnectionString)
    $accessConnection.Open()
    
    # Get the schema information for the tables in the MS Access database
    $accessTables = $accessConnection.GetSchema("Tables")

    foreach ($accessTable in $accessTables.Rows) {
        $tableName = $accessTable["TABLE_NAME"]
        
        # Skip system tables
        if ($tableName -like "MSys*") {
            continue
        }
        
        Write-Output "Checking table: $tableName"
        
        # Get column count for Access table
        $accessColumns = $accessConnection.GetSchema("Columns") | Where-Object { $_["TABLE_NAME"] -eq $tableName }
        $accessColumnCount = $accessColumns.Count
        
        # Get column count for SQL Server table
        $sqlColumnQuery = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '$tableName' AND COLUMN_NAME <> 'TempInvestigationID'"
        $sqlColumnCommand = New-Object System.Data.SqlClient.SqlCommand($sqlColumnQuery, $sqlConnection)
        $sqlColumnCount = $sqlColumnCommand.ExecuteScalar()
        
        Write-Output "Column count - Access: $accessColumnCount, SQL Server: $sqlColumnCount"
        
        if ($accessColumnCount -ne $sqlColumnCount) {
            Write-Output "WARNING: Column count mismatch for table $tableName!"
        }
        
        # Get row count for Access table
        $accessRowQuery = "SELECT COUNT(*) FROM [$tableName]"
        $accessRowCommand = New-Object System.Data.OleDb.OleDbCommand($accessRowQuery, $accessConnection)
        $accessRowCount = $accessRowCommand.ExecuteScalar()
        
        # Get row count for SQL Server table
        $sqlRowQuery = "SELECT COUNT(*) FROM [$tableName]"
        $sqlRowCommand = New-Object System.Data.SqlClient.SqlCommand($sqlRowQuery, $sqlConnection)
        $sqlRowCount = $sqlRowCommand.ExecuteScalar()
        
        Write-Output "Row count - Access: $accessRowCount, SQL Server: $sqlRowCount"
        
        if ($accessRowCount -ne $sqlRowCount) {
            Write-Output "WARNING: Row count mismatch for table $tableName!"
        }
    }
    
    # Close the Access connection
    $accessConnection.Close()
}

# Close SQL Server connection
$sqlConnection.Close()
