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
            $data = Extract-DataFromAccess -dbPath $dbPath -tableName $tableName
            # Add the DatabaseID column with the folder name as its value
            $data.Columns.Add("DatabaseID", [System.String]) | Out-Null
            foreach ($row in $data.Rows) {
                $row["DatabaseID"] = $dbFolderName
            }
            Write-DataToSQL -dataTable $data -tableName $tableName -sqlConnStr $sqlConnStr
        } catch {
            Write-Output "Error processing table $tableName in database $dbPath: $_"
        }
    }
}

Write-Output "Data consolidation to SQL Server is complete."