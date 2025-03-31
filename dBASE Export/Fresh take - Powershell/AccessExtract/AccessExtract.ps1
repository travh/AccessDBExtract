# Define the path to the MS Access database
$databasePath = "C:\AccessTest\DB1\Database.accdb"

# Create a connection string
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$databasePath;"

# Create a new OleDbConnection object
$connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)

# Open the connection
$connection.Open()

# Get the schema information for the tables
$tables = $connection.GetSchema("Tables")

# Loop through each table and get the columns
foreach ($table in $tables.Rows) {
    $tableName = $table["TABLE_NAME"]
    Write-Output "Table: $tableName"

    # Get the schema information for the columns in the table
    $columns = $connection.GetSchema("Columns", @($null, $null, $tableName, $null))

    foreach ($column in $columns.Rows) {
        $columnName = $column["COLUMN_NAME"]
        Write-Output "    Column: $columnName"
    }
}

# Close the connection
$connection.Close()