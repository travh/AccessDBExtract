# Define the root folder containing MS Access databases
$accessRootFolder = "C:\AccessTest"  # Change this to your root folder

# Define the connection string for SQL Server
$sqlServerInstance = "DESKTOP-VKMSDNG"
$sqlDatabaseName = "AccessStaging"
$sqlConnectionString = "Server=$sqlServerInstance;Database=$sqlDatabaseName;Integrated Security=True;"

# Create a new SqlConnection object for SQL Server
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection($sqlConnectionString)
$sqlConnection.Open()

# Get all Access database files in the root folder and subfolders
$accessDatabaseFiles = Get-ChildItem -Path $accessRootFolder -Recurse -Filter "*.accdb"

foreach ($accessDatabase in $accessDatabaseFiles) {
    $sourceFolderName = Split-Path -Path $accessDatabase.DirectoryName -Leaf
    Write-Output "Processing database: $($accessDatabase.FullName) from folder: $sourceFolderName"
    
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
        
        Write-Output "Copying data from $tableName..."
        
        # Read data from Access table
        $selectQuery = "SELECT * FROM [$tableName]"
        $accessCommand = New-Object System.Data.OleDb.OleDbCommand($selectQuery, $accessConnection)
        $dataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter($accessCommand)
        $dataTable = New-Object System.Data.DataTable
        $dataAdapter.Fill($dataTable) | Out-Null
        
        if ($dataTable.Rows.Count -eq 0) {
            Write-Output "No data found in $tableName. Skipping..."
            continue
        }
        
        # Insert data into SQL Server
        foreach ($row in $dataTable.Rows) {
            $columnNames = ($dataTable.Columns | ForEach-Object { "[$($_.ColumnName)]" }) -join ", "
            $values = ($dataTable.Columns | ForEach-Object { "'" + $row[$_.ColumnName].ToString().Replace("'", "''") + "'" }) -join ", "
            
            # Include TempInvestigationID column
            $columnNames += ", [TempInvestigationID]"
            $values += ", '$sourceFolderName'"
            
            $insertQuery = "INSERT INTO [$tableName] ($columnNames) VALUES ($values)"
            
            try {
                $sqlCommand = New-Object System.Data.SqlClient.SqlCommand($insertQuery, $sqlConnection)
                $sqlCommand.ExecuteNonQuery() | Out-Null
            } catch {
                Write-Output "Error inserting into ${tableName}: $($_.Exception.Message)"
            }
        }
        
        Write-Output "Data copied for $tableName."
    }
    
    # Close the Access connection
    $accessConnection.Close()
}

# Close SQL Server connection
$sqlConnection.Close()
