# Alexander Barker 012100649

# Set Debug Bool
$DebugPreference = "SilentlyContinue"

# Define the database name
$dbName = "ClientDB"
# Define the SQL Server instance name
$sqlInstanceName = "localhost\SQLEXPRESS"
# Define the table name to be created
$tableName = "Client_A_Contacts"

try {
    # Connection string to SQL Server with TrustServerCertificate and Windows Authentication
    $connectionString = "Server=$sqlInstanceName;Database=master;Trusted_Connection=True;TrustServerCertificate=True;"

    # Check if the database exists, then drop it if it does
    Write-Debug "Checking if the database $dbName exists"
    $dbExistsQuery = "SELECT name FROM sys.databases WHERE name = '$dbName'"
    $dbExists = Invoke-Sqlcmd -ConnectionString $connectionString -Query $dbExistsQuery

    if ($dbExists) {
        Write-Debug "$dbName exists and will be dropped"
        Write-Host "$dbName exists and will be dropped."

        # Prepare the SQL query to drop the database
        Write-Debug "Preparing the SQL query to drop the database"
        $dropQuery = "ALTER DATABASE [$dbName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [$dbName]"

        # Execute the SQL command to drop the database
        Write-Debug "Executing the SQL command to drop the database"
        Invoke-Sqlcmd -ConnectionString $connectionString -Query $dropQuery
        Write-Host "$dbName has been dropped"
    } else {
        Write-Debug "$dbName does not exist; preparing to create it"
        Write-Host "$dbName does not exist. Proceeding to create the database"
    }

    # Create SQL database 
    Write-Debug "Creating the database $dbName using a SQL command"
    Write-Host "Creating the database $dbName using SQL command"
    $createDbQuery = "CREATE DATABASE [$dbName]"

    # Create the database
    Write-Debug "Executing the SQL command to create the database"
    Invoke-Sqlcmd -ConnectionString $connectionString -Query $createDbQuery
    Write-Host "$dbName database was successfully created."

    # Update the connection string to use the new database
    $connectionString = "Server=$sqlInstanceName;Database=$dbName;Trusted_Connection=True;TrustServerCertificate=True;"

    # Create the table using a direct SQL command
    Write-Debug "Preparing to create the table $tableName"
    Write-Host "Creating the table $tableName"
    $createTableQuery = @"
    CREATE TABLE [$tableName] (
        first_name NVARCHAR(50),
        last_name NVARCHAR(50),
        city NVARCHAR(50),
        county NVARCHAR(50),
        zip NVARCHAR(10),
        officePhone NVARCHAR(15),
        mobilePhone NVARCHAR(15)
    )
"@
    # Command to create table
    Write-Debug "Executing the SQL command to create the table"
    Invoke-Sqlcmd -ConnectionString $connectionString -Query $createTableQuery
    Write-Host "$tableName table was successfully created."

    # Import data from the CSV file
    Write-Debug "Importing leads data from CSV file"
    Write-Host "Importing leads data from CSV file"
    $Leads = Import-Csv .\NewClientData.csv

    # Get the count of total records
    Write-Debug "Counting the total number of records to be inserted"
    $AllUsers = $Leads.count
    $count = 1

    # Base insert SQL statement
    Write-Debug "Preparing the base SQL insert statement"
    $Insert = "INSERT INTO [$($tableName)] (first_name, last_name, city, county, zip, officePhone, mobilePhone)"

    Write-Debug "Starting to insert records into $tableName"
    Write-Host "Starting to insert records into $tableName"

    # Loop through records in the CSV
    ForEach ($u in $Leads) {
        Write-Debug "Inserting record $count`: $($u.first_name) $($u.last_name)"

        # Display progress to the user
        $progress = "Adding new SQL user $($u.first_name) $($u.last_name). $($count) of $($AllUsers)"
        Write-Progress -Activity "Restore SQL DB" -Status $progress -PercentComplete (($count / $AllUsers) * 100)

        # Prepare the values for the insert statement
        Write-Debug "Preparing the values for the insert statement"
        $Values = "VALUES ( `
                    '$($u.first_name)', `
                    '$($u.last_name)', `
                    '$($u.city)', `
                    '$($u.county)', `
                    '$($u.zip)', `
                    '$($u.officePhone)', `
                    '$($u.mobilePhone)')"

        # Combine the base insert statement and values
        Write-Debug "Executing the insert statement for record $count"
        $query = $Insert + $Values
        Invoke-Sqlcmd -ConnectionString $connectionString -Query $query

        # Count for progress tracking
        Write-Debug "Incrementing the record count"
        $count++
    }

    # Export to text file for review
    Write-Debug "Exporting the table data to SqlResults.txt"
    Write-Host "Exporting table data to SqlResults.txt"
    Invoke-Sqlcmd -ConnectionString $connectionString -Query "SELECT * FROM dbo.$($tableName)" | Out-File -FilePath .\SqlResults.txt
    Write-Host "$dbName database has been built with no errors! SqlResults file has been created."
}
catch {
    # Catch and display any errors
    Write-Debug "An error occurred: $($_.Exception.Message)"
    Write-Host "Something went wrong"
    Write-Host $_
}
