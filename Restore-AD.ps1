# Alexander Barker 012100649

# Debug Bool
$DebugPreference = "SilentlyContinue"

# Set AD Root
$RootAD = (Get-ADDomain).distinguishedName

# Set OU Canonical Name
$OUCName = "Finance"

# Set OU Display Name
$OUDName = "Finance"

# Set Path to the OU
$ADPath = "OU=$($OUCName),$($RootAD)"

# Import users from a CSV file 
Write-Debug "Importing user data from CSV file"
$FinUsers = Import-Csv -Path ".\financePersonnel.csv"

try {
    # Verify if the OU 'Finance' exists
    Write-Debug "Checking if the OU $OUCName exists in Active Directory"
    Write-Host "Starting script"

    if (-Not([ADSI]::Exists("LDAP://$($ADPath)"))) {
        # Create the OU if it does not exist
        Write-Debug "$OUCName does not exist. Creating the OU"
        Write-Host "$($OUCName) does not exist. Creating OU"
        New-ADOrganizationalUnit -Name $OUCName -Path $RootAD -DisplayName $OUDName -ProtectedFromAccidentalDeletion $false
        Write-Host "$($OUCName) has been created!"
        Write-Debug "OU $OUCName has been created"

        # Count the number of users to be added
        Write-Debug "Counting the number of users to be added"
        $UserList = $FinUsers.count
        $count = 1
        
        # Iterate over each user and add them to AD
        Write-Debug "Starting to add users to the $OUCName OU"
        ForEach ($fu in $FinUsers) {
            Write-Debug "Adding user $count: $($fu.First_Name) $($fu.Last_Name)"
            $First = $fu.First_Name
            $Last = $fu.Last_Name
            $Name = "$($First) $($Last)"
            $SamAcct = $fu.samAccount
            $Zip = $fu.PostalCode
            $WorkPhone = $fu.OfficePhone
            $Mobile = $fu.MobilePhone

            # Add the user to AD
            Write-Debug "Creating AD user: $Name"
            New-ADUser -GivenName $First `
            -Surname $Last `
            -Name $Name `
            -SamAccountName $SamAcct `
            -DisplayName $Name `
            -PostalCode $Zip `
            -OfficePhone $WorkPhone `
            -MobilePhone $Mobile `
            -Path $ADPath

            # Increment the counter
            Write-Debug "Incrementing user count"
            $count++
        }
    } else {
        # If the OU exists remove it
        Write-Debug "OU $OUCName already exists. Preparing to remove it"
        Write-Host "$($OUCName) exists. Removing $($OUCName)"
        Remove-ADOrganizationalUnit -Identity $ADPath -Confirm:$false -Recursive
        Write-Host "$($OUCName) removed. Run script again."
        Write-Debug "OU $OUCName has been removed"
        Exit
    }

    # Indicate that the Active Directory task completed successfully
    Write-Debug "Active Directory task completed successfully"
    Write-Host "Restore Active Directory Task Completed with no errors."

    # Output the results to a text file
    Write-Debug "Writing AD user results to AdResults.txt"
    Write-Host "Writing ADResults.txt"
    Get-ADUser -Filter * -SearchBase $ADPath -Properties DisplayName, PostalCode, OfficePhone, MobilePhone > .\AdResults.txt
}
catch {
    # Catch and display any errors
    Write-Debug "An error occurred: $($_.Exception.Message)"
    Write-Host "Error occurred:"
    Write-Host $_
}
