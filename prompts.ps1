# Alexander Barker 012100649

# Create a “switch” statement that continues to prompt a user by doing each of the following activities, until a user presses key 5

$DebugPreference = "SilentlyContinue"

do {
    Write-Host "`nSelect an option:"
    Write-Host "1. Produce Daily Log File"
    Write-Host "2. List the file"
    Write-Host "3. Display CPU and RAM usage"
    Write-Host "4. Display running processes"
    Write-Host "5. Exit"
    
    # Prompt for user input
    $userInput = Read-Host "Enter your choice (1-5)"
    
    # Switch statement to handle user input
    switch ($userInput) {
        1 {
            try {

                # Using a regular expression, list files within the Requirements1 folder, with the .log file extension 
                #and redirect the results to a new file called “DailyLog.txt” within the same directory without overwriting existing data. 
                # Each time the user selects this prompt, the current date should precede the listing. (User presses key 1.)

                Write-Debug "Producing Daily Log File"
                # Get Current Date
                $CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Write-Debug "Current Date $CurrentDate."

                $logFilePath = ".\DailyLog.txt"
                Write-Debug "Appending Date to DailyLog.txt"
                Add-Content -Path $logFilePath -Value "$currentDate"

                Write-Debug "Searching For .log Files & Appending Log Files"
                Get-ChildItem -Path . -File | Where-Object { $_.Name -match '\.log$' } | ForEach-Object {
                    Add-Content -Path $logFilePath -Value $_.FullName
                }

                Write-Host "`nLog files have been appended to DailyLog.txt"
            } catch [System.OutOfMemoryException] {
                Write-Host "An out-of-memory error occurred while producing the daily log file."
                Write-Debug "OutOfMemoryException caught in option 1."
            } catch {
                Write-Host "An unexpected error occurred: $($_.Exception.Message)"
            }
        }
        2 {
            try {
                
                #List the files inside the “Requirements1” folder in tabular format, sorted in ascending alphabetical order. 
                #Direct the output into a new file called “C916contents.txt” found in your “Requirements1” folder. (User presses key 2.)

                # Get-ChildItem -Path . -File - Gets all items in our current directory.
                # Sort-Object Name = Sort all Files by name
                # Format-Table Mode, LastWriteTime, Length, Name -AutoSize | Out-String | Add-Content -Path "./C916contents.txt" = Format Table returns output that needs to be turned into a string again after then print to file

                Write-Debug "Performing Enum in Directory, Sorting and Appending."
                Get-ChildItem -Path . -File | Sort-Object Name | Format-Table Mode, LastWriteTime, Length, Name -AutoSize | Out-String | Add-Content -Path "./C916contents.txt"
                Write-Debug "`nC916contents.txt Created"
            } catch [System.OutOfMemoryException] {
                Write-Host "An out-of-memory error occurred while listing and formatting files."
                Write-Debug "OutOfMemoryException caught in option 2."
            } catch {
                Write-Host "An unexpected error occurred: $($_.Exception.Message)"
            }
        }
        3 {
            try {
                Write-Debug "Retrieving performance counter data and assigning it to variable"
                # Get the total CPU usage percentage To do this we count logical processors
                $CpuUsageSample = Get-Counter '\Processor(_Total)\% Processor Time'
                Write-Debug "Variable value: $CpuUsageSample"

                 # Extract the current CPU usage value and round it to 2 decimal places
                $TotalCpuUsage = [Decimal]::Round($CpuUsageSample.CounterSamples[0].CookedValue, 2)
                Write-Debug "Total Usage: $TotalCpuUsage"

                # Calculate the free CPU percentage
                $FreeCpuPercentage = [Decimal]::Round(100 - $TotalCpuUsage, 2)
                Write-Debug "Free Percentage: $FreeCpuPercentage"

                # Print results to terminal
                Write-Host "Used CPU Percentage: $TotalCpuUsage%"
                Write-Host "Free CPU Percentage: $FreeCpuPercentage%"

                # Get memory usage
                Write-Debug "Getting Memory information"
                $memory = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object TotalVisibleMemorySize, FreePhysicalMemory
                Write-Debug "Memory: $memory"
                Write-Debug "Getting Total Memory"
                $totalMemory = [math]::round($memory.TotalVisibleMemorySize / 1MB, 2)
                Write-Debug "Total Memory: $totalMemory"
                Write-Debug "Getting Total Free Memory"
                $freeMemory = [math]::round($memory.FreePhysicalMemory / 1MB, 2)
                Write-Debug "Total Free Memory: $freeMemory"
                Write-Debug "Calculating Used Memory"
                $usedMemory = [math]::round($totalMemory - $freeMemory, 2)
                Write-Debug "Total Used Memory: $usedMemory"

                Write-Debug "Calculating Total Usage"
                $usagePercentage = [math]::round(($usedMemory / $totalMemory) * 100, 2)
                Write-Debug "Total Usage: $usagePercentage"

                Write-Host "`nMemory Usage:"
                Write-Host "Total Memory: $totalMemory MB"
                Write-Host "Used Memory: $usedMemory MB"
                Write-Host "Free Memory: $freeMemory MB"
                Write-Host "Usage: $usagePercentage`%"
            } catch [System.OutOfMemoryException] {
                Write-Host "An out-of-memory error occurred while calculating CPU and RAM usage."
                Write-Debug "OutOfMemoryException caught in option 3."
            } catch {
                Write-Host "An unexpected error occurred: $($_.Exception.Message)"
            }
        }
        4 {
            try {

                # Get all running processes, convert to 64bit to avoid int overflow. Round and then push processes to grid view in ascending order
                Get-Process | Sort-Object VirtualMemorySize64 | 
                Select-Object Id, ProcessName, @{Name="VirtualMemorySize (MB)"; Expression={[math]::round($_.VirtualMemorySize64 / 1MB, 2)}} | 
                Out-GridView -Title "Processes Sorted by Virtual Memory Size"
            } catch [System.OutOfMemoryException] {
                Write-Host "An out-of-memory error occurred while displaying running processes."
                Write-Debug "OutOfMemoryException caught in option 4."
            } catch {
                Write-Host "An unexpected error occurred: $($_.Exception.Message)"
            }
        }
        5 {
            Write-Host "Exiting."
            break
        }
        Default {
            Write-Host "Please enter a number between 1 and 5."
        }
    }
} while ($userInput -ne "5")
