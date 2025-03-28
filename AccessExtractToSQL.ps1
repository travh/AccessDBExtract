# Suppress specific UserWarning
[System.Reflection.Assembly]::LoadWithPartialName("System.Data.Odbc") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Data.SqlClient") | Out-Null

# Function to get all table names from MS Access database
function Get-TableNames {
    param (
        [string]$dbPath
    )
    $connStr = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=$dbPath"
    $conn = New-Object System.Data.Odbc.OdbcConnection($connStr)
    $conn.Open()
    $tables = $conn.GetSchema("Tables")
    $tableNames = $tables | Where-Object { $_["TABLE_TYPE"] -eq "TABLE" } | Select-Object -ExpandProperty TABLE_NAME
    $conn.Close()
    return $tableNames
}

# Function to read data from a specific table in MS Access database
function Get-DataFromAccess {
    param (
        [string]$dbPath,
        [string]$tableName
    )
    $connStr = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=$dbPath"
    $conn = New-Object System.Data.Odbc.OdbcConnection($connStr)
    $conn.Open()
    $query = "SELECT * FROM [$tableName]"
    $cmd = New-Object System.Data.Odbc.OdbcCommand($query, $conn)
    $adapter = New-Object System.Data.Odbc.OdbcDataAdapter($cmd)
    $dataSet = New-Object System.Data.DataSet
    $adapter.Fill($dataSet)
    $conn.Close()
    return $dataSet.Tables
}

# Function to write data to SQL Server database
function Write-DataToSQL {
    param (
        [System.Data.DataTable]$dataTable,
        [string]$tableName,
        [string]$sqlConnStr
    )
    try {
        $bulkCopy = New-Object Data.SqlClient.SqlBulkCopy($sqlConnStr)
        $bulkCopy.DestinationTableName = $tableName
        $bulkCopy.WriteToServer($dataTable)
        Write-Output "Successfully wrote data to SQL Server table $tableName"
    } catch {
        Write-Output ("Error writing to SQL Server table {0}: {1}" -f $tableName, $_.Exception.Message)
        throw
    }
}

# Function to traverse directories and find all Access databases
function Find-AccessDatabases {
    param (
        [string]$folder
    )
    $accessDatabases = Get-ChildItem -Path $folder -Recurse -Filter *.accdb | Select-Object -ExpandProperty FullName
    return $accessDatabases
}

# Example usage
$dbFolder = "C:\AccessTest"
$sqlConnStr = "Server=DESKTOP-VKMSDNG;Database=AccessStaging;Integrated Security=True;"

$accessDatabases = Find-AccessDatabases -folder $dbFolder

$totalDatabases = $accessDatabases.Count
for ($i = 0; $i -lt $totalDatabases; $i++) {
    Write-Output "Extracting from database $($i + 1) of $totalDatabases"
    $dbPath = $accessDatabases[$i]
    $dbFolderName = Split-Path -Path $dbPath -Leaf
    $tableNames = Get-TableNames -dbPath $dbPath
    foreach ($tableName in $tableNames) {
        try {
            Write-Output "Processing table $tableName in database $dbPath"
            $data = Get-DataFromAccess -dbPath $dbPath -tableName $tableName
            # Create a new DataTable with the additional DatabaseID column
            $newDataTable = New-Object System.Data.DataTable
            $data.Columns | ForEach-Object { $newDataTable.Columns.Add($_.ColumnName, $_.DataType) }
            $newDataTable.Columns.Add("DatabaseID", [System.String])
            # Copy data to the new DataTable
            foreach ($row in $data.Rows) {
                $newRow = $newDataTable.NewRow()
                $newRow.ItemArray = $row.ItemArray
                $newRow["DatabaseID"] = $dbFolderName
                $newDataTable.Rows.Add($newRow)
            }
            Write-DataToSQL -dataTable $newDataTable -tableName $tableName -sqlConnStr $sqlConnStr
        } catch {
            Write-Output ("Error processing table {0} in database {1}: {2}" -f $tableName, $dbPath, $_.Exception.Message)
        }
    }
}

Write-Output "Data consolidation to SQL Server is complete."