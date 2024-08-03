# Specify the OU path and group name
$ouPath = "OU=placeholder,OU=placeholder,DC=placeholder,DC=placeholder"  # Replace with the desired OU path
$groupName = "NET_VPN_EMPLOYEE"  # Replace with the desired group name

# Connect to the Active Directory module
Import-Module ActiveDirectory

# Get the users in the specified OU
$users = Get-ADUser -Filter {(EmployeeID -ne "NA")} -SearchBase $ouPath

# Initialize an array to store the results
$results = @()

# Iterate through each user and check if they belong to the specified group
foreach ($user in $users) {
    $groupMembership = Get-ADUser $user -Properties MemberOf |
                      Select-Object -ExpandProperty MemberOf

    if ($groupMembership -contains $groupName) {
        $result = @{
            Username = $user.SamAccountName
            Group = $groupName
        }
        $results += $result
    }
}

# Specify the output file path
$outputFilePath = "C:\VPN_Users.txt"  # Replace with the desired output file path

# Export the results to a text file
$results | Export-Csv -Path $outputFilePath -NoTypeInformation
