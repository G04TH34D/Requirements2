# Alexander Barker 012100649


#A.  Create a PowerShell script named “Restore-AD.ps1” within the attached “Requirements2” folder. Create a comment block and include your first and last name along with your student ID.
#B.  Write the PowerShell commands in “Restore-AD.ps1” that perform all the following functions without user interaction:
#        1.  Check for the existence of an Active Directory Organizational Unit (OU) named “Finance.” 
#        Output a message to the console that indicates if the OU exists or if it does not. 
#        If it already exists, delete it and output a message to the console that it was deleted.
#        2.  Create an OU named “Finance.” Output a message to the console that it was created.
#        3.  Import the financePersonnel.csv file (found in the attached “Requirements2” directory) into your Active Directory domain and directly into the finance OU.
#        Be sure to include the following properties:
                #•   First Name
                #•   Last Name
                #•   Display Name (First Name + Last Name, including a space between)
                #•   Postal Code
                #•   Office Phone
                #•   Mobile Phone
#        4.  Include this line at the end of your script to generate an output file for submission:
            # Get-ADUser -Filter * -SearchBase “ou=Finance,dc=consultingfirm,dc=com” -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\AdResults.txt


# Debug Bool
$DebugPreference = "SilentlyContinue"
# Set AD Root
$RootAD = (Get-ADDomain).distinguishedName
# Set OU Canonl Name
$OUCName = "Finance"
# Set OU Display Name
$OUDName = "Finance"
# Set Path to the OU
$ADPath = "OU=$($OUCName),$($RootAD)"
# Import users from a CSV file 
Write-Debug "Importing user data from CSV file"
$FinUsers = Import-Csv -Path ".\financePersonnel.csv"


try {

    # 1.  Check for the existence of an Active Directory Organizational Unit (OU) named “Finance.” 
    # Output a message to the console that indicates if the OU exists or if it does not. 
    # If it already exists, delete it and output a message to the console that it was deleted.

    # Verify if the OU 'Finance' exists
    Write-Debug "Starting script"
    Write-Debug "Checking if the OU $OUCName exists"
    Write-Host "Starting script"
    Write-Host "Checking if the OU $OUCName exists"


    # If it already exists,
    if ([ADSI]::Exists("LDAP://$($ADPath)")) {

        # Output a message to the console that it was deleted
        Write-Debug "OU $OUCName exists removing it"
        Write-Host "$($OUCName) exists removing $($OUCName)"
        Write-Debug "Attempting to remove $($OUCName)"

        # Delete it 
        Remove-ADOrganizationalUnit -Identity $ADPath -Confirm:$false -Recursive
        # Output a message to the console that it was deleted
        Write-Debug "OU $OUCName has been removed"
        Write-Host "OU $OUCName has been removed"

        # 2.  Create an OU named “Finance.”
        Write-Debug "Creating the OU $OUCName"
        Write-Host "Creating OU $($OUCName) "
        New-ADOrganizationalUnit -Name $OUCName -Path $RootAD -DisplayName $OUDName -ProtectedFromAccidentalDeletion $false
        # Output a message to the console that it was created.
        Write-Debug "OU $OUCName has been created"
        Write-Host "$($OUCName) has been created"

        # Progress and Debugging
        Write-Debug "Counting the number of users to be added"
        $UserList = $FinUsers.count
        $count = 1
        
        # Iterate over each user and add them to AD
        Write-Debug "Starting to add users to the $OUCName OU"
        ForEach ($fu in $FinUsers) {
            #Write-Debug "Adding user $count`: $($fu.First_Name) $($fu.Last_Name)"
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

    # Indicate task completed successfully
    Write-Debug "Task completed successfully"
    Write-Host "Restore Task Completed with no errors."

    # Output the results to a text file
    Write-Debug "Writing results to AdResults.txt"
    Write-Host "Writing ADResults.txt"
    Get-ADUser -Filter * -SearchBase $ADPath -Properties DisplayName, PostalCode, OfficePhone, MobilePhone > .\AdResults.txt

    } else {
        
        # Output a message to the console that indicates if the OU exists or if it does not.
        Write-Debug "OU $OUCName does not exist"
        Write-Host "OU $OUCName does not exist"
           
        # 2.  Create an OU named “Finance.”
        Write-Debug "Creating the OU $OUCName"
        Write-Host "Creating OU $($OUCName) "
        New-ADOrganizationalUnit -Name $OUCName -Path $RootAD -DisplayName $OUDName -ProtectedFromAccidentalDeletion $false
       
        # Output a message to the console that it was created.
        Write-Debug "OU $OUCName has been created"
        Write-Host "$($OUCName) has been created"

        # Progress and Debugging
        Write-Debug "Counting the number of users to be added"
        $UserList = $FinUsers.count
        $count = 1
        
        # Iterate over each user and add them to AD
        Write-Debug "Starting to add users to the $OUCName OU"
        ForEach ($fu in $FinUsers) {
            #Write-Debug "Adding user $count`: $($fu.First_Name) $($fu.Last_Name)"
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

    # Indicate task completed successfully
    Write-Debug "Task completed successfully"
    Write-Host "Restore Task Completed with no errors."

    # Output the results to a text file
    Write-Debug "Writing results to AdResults.txt"
    Write-Host "Writing ADResults.txt"
    Get-ADUser -Filter * -SearchBase $ADPath -Properties DisplayName, PostalCode, OfficePhone, MobilePhone > .\AdResults.txt
    }
}
catch {
    # Catch and display any errors
    Write-Debug "An error occurred: $($_.Exception.Message)"
    Write-Host "Error occurred:"
    Write-Host $_
}