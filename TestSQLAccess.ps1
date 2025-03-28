# Define the SQL Server connection string
$sqlConnStr = "Server=DESKTOP-VKMSDNG;Database=AccessStaging;Integrated Security=True;"

# Function to test SQL Server connection and retrieve table names
function Test-SQLServerConnection {
    param (
        [string]$sqlConnStr
    )
    try {
        # Create a new SQL connection
        $conn = New-Object System.Data.SqlClient.SqlConnection($sqlConnStr)
        $conn.Open()
        
        # Define the query to retrieve table names
        $query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
        
        # Execute the query
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $query
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dataSet = New-Object System.Data.DataSet
        $adapter.Fill($dataSet)
        
        # Close the connection
        $conn.Close()
        
        # Output the table names
        $dataSet.Tables | ForEach-Object { Write-Output $_.TABLE_NAME }
        
        Write-Output "Connection to SQL Server and database access verified successfully."
    } catch {
        Write-Output ("Error connecting to SQL Server: {0}" -f $_.Exception.Message)
    }
}

# Test the SQL Server connection
Test-SQLServerConnection -sqlConnStr $sqlConnStr