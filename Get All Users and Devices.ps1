# Connect-MgGraph

# Retrieve users from department '182'
$allFieldUsers = Get-MgUser -count userCount -ConsistencyLevel eventual -Filter "startsWith(department, '182')" -All | Select-Object -ExpandProperty UserPrincipalName

# Initialize an array to hold the results
$results = @()

try {
    # Iterate over each user
    foreach ($fieldUser in $allFieldUsers) {
        $user = Get-MgUser -UserId $fieldUser
        $jobTitle = $user.JobTitle

        # Retrieve the devices owned by the user
        $userDevices = Get-MgUserOwnedDeviceAsDevice -UserId $fieldUser | Select-Object -ExpandProperty DisplayName

        # Iterate over each device
        foreach ($userDevice in $userDevices) {
            # Check if the device name starts with "C-" or "N-*"
            if ($userDevice -like "C-*" -or $userDevice -like "N-*") {
                # Create a custom object with the required fields
                $result = [PSCustomObject]@{
                    JobTitle    = $jobTitle
                }
                # Add the custom object to the results array
                $results += $result
            }
        }
    }
}
catch {
    # Handle any errors
    Write-Host "ERROR: $_" -ForegroundColor Red
}

# Export the results to a CSV file
$uniqueTitles = $results | Sort-Object
$uniqueTitles | Export-Csv -Path "C:\Script\UniqueFieldUsersDevices.csv" -NoTypeInformation