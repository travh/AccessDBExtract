function Export-TablesFromAccessDB {
    param (
        [string]$dbPath,
        [string]$exportPath,
        [string]$folderName
    )

    $accessApp = New-Object -ComObject Access.Application
    $accessApp.OpenCurrentDatabase($dbPath)

    $tables = $accessApp.CurrentDb.TableDefs | Where-Object { $_.Name -notmatch "^(MSys|~)" }

    foreach ($table in $tables) {
        $tableName = $table.Name
        $exportFile = Join-Path $exportPath "$folderName`_$tableName.dbf"
        $accessApp.DoCmd.TransferDatabase(1, "dBase IV", $exportPath, 0, $tableName, "$folderName`_$tableName.dbf")
        Write-Output "Exported $tableName to $exportFile"
    }

    $accessApp.CloseCurrentDatabase()
    $accessApp.Quit()
}

function Export-AllAccessDBs {
    param (
        [string]$rootFolder,
        [string]$exportPath
    )

    $accessDBs = Get-ChildItem -Path $rootFolder -Recurse -Filter *.accdb

    foreach ($db in $accessDBs) {
        $folderName = Split-Path -Leaf (Split-Path -Parent $db.FullName)
        Write-Output "Processing $($db.FullName)"
        Export-TablesFromAccessDB -dbPath $db.FullName -exportPath $exportPath -folderName $folderName
    }
}

# Set the root folder to search for Access databases and the export path
$rootFolder = "C:\AccessTest\"
$exportPath = "C:\AccessTest\Output\dBASE"

# Run the export process
Export-AllAccessDBs -rootFolder $rootFolder -exportPath $exportPath
