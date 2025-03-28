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
    $bulkCopy = New-Object Data.SqlClient.SqlBulkCopy($sqlConnStr)
    $bulkCopy.DestinationTableName = $tableName
    $bulkCopy.WriteToServer($dataTable)
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
            $data = Get-DataFromAccess -dbPath $dbPath -tableName $tableName
            # Add the DatabaseID column with the folder name as its value
            $data.Columns.Add("DatabaseID", [System.String]) | Out-Null
            foreach ($row in $data.Rows) {
                $row["DatabaseID"] = $dbFolderName
            }
            Write-DataToSQL -dataTable $data -tableName $tableName -sqlConnStr $sqlConnStr
        } catch {
            Write-Output ("Error processing table {0} in database {1}: {2}" -f $tableName, $dbPath, $_.Exception.Message)
        }
    }
}

Write-Output "Data consolidation to SQL Server is complete."