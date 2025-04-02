# Define the root folder and filename to search for
$searchRootFolder = "C:\AccessTest"  # Change this to your target folder
$targetFileName = "database.accdb"  # Name of the file to locate
$csvOutputFolder = "C:\AccessTest\SchemaOutput"  # Change this to your desired CSV output folder

# Get all matching files in the root folder and subfolders
$matchingFiles = Get-ChildItem -Path $searchRootFolder -Recurse | Where-Object { $_.Name -eq $targetFileName }

if ($matchingFiles.Count -eq 0) {
    Write-Output "No files found matching '$targetFileName' in '$searchRootFolder'."
    exit
}

# Determine the earliest and latest creation dates
$earliestFile = $matchingFiles | Sort-Object CreationTime | Select-Object -First 1
$latestFile = $matchingFiles | Sort-Object CreationTime -Descending | Select-Object -First 1

Write-Output "Earliest file: $($earliestFile.FullName) - Created on: $($earliestFile.CreationTime)"
Write-Output "Latest file: $($latestFile.FullName) - Created on: $($latestFile.CreationTime)"

# Function to map Access data type numbers to names
function Get-AccessDataTypeName {
    param ($dataTypeNumber)
    
    switch ($dataTypeNumber) {
        2 { return "SHORT" }
        3 { return "LONG" }
        4 { return "SINGLE" }
        5 { return "DOUBLE" }
        6 { return "CURRENCY" }
        7 { return "DATE/TIME" }
        11 { return "YESNO" }
        17 { return "BYTE" }
        130 { return "TEXT" }
        203 { return "MEMO" }
        default { return "UNKNOWN" }
    }
}

# Function to extract schema information from an Access database
function Get-AccessSchema {
    param ($databasePath)
    
    $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$databasePath;"
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $connection.Open()
    
    $schemaTable = $connection.GetSchema("Columns")
    $connection.Close()
    
    return $schemaTable | Select-Object TABLE_NAME, COLUMN_NAME, @{Name='DATA_TYPE'; Expression={ Get-AccessDataTypeName $_.DATA_TYPE }}
}

# Extract schema information from earliest and latest files
$earliestSchema = Get-AccessSchema -databasePath $earliestFile.FullName
$latestSchema = Get-AccessSchema -databasePath $latestFile.FullName

# Ensure the output directory exists
if (!(Test-Path -Path $csvOutputFolder)) {
    New-Item -ItemType Directory -Path $csvOutputFolder | Out-Null
}

# Output schema to CSV files
$earliestSchema | Export-Csv -Path "$csvOutputFolder\Earliest_Schema.csv" -NoTypeInformation
$latestSchema | Export-Csv -Path "$csvOutputFolder\Latest_Schema.csv" -NoTypeInformation

Write-Output "Schema information exported to $csvOutputFolder\Earliest_Schema.csv and $csvOutputFolder\Latest_Schema.csv"
