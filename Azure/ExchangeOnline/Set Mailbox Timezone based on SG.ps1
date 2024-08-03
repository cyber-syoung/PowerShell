$WarningPreference = 'SilentlyContinue'
# Connect-MgGraph
# Connect-ExchangeOnline

$script:groups = @{
    "83bc8984-9ef4-4912-ad74-d5833b92df7d" = "Pacific Standard Time" # "TIMEZONE: Pacific Standard Time (PST)"
    "265d1cd2-5d01-4226-9143-cc9db8c73462" = "Mountain Standard Time" # "TIMEZONE: Mountain Standard Time (MST)"
    "500a1dff-044a-4765-b525-6d7603349267" = "Central Standard Time" # "TIMEZONE: Central Standard Time (CST)"
    "594dad4b-c389-43ce-a3e5-9bbc51bf6edd" = "Eastern Standard Time" # "TIMEZONE: Eastern Standard Time (EST)"
}

# DEBUG variable
$script:allUsers = @("samuel.young@confidential.com")

# PROD variable
#$script:allUsers = Get-MgUser -All | Select-Object -ExpandProperty UserPrincipalName 

## TODO: Create a function to check if the user's in the specific group, and confirm that they have the correct time zone applied. 
## TODO: Update Set-TimeZone to include the above TODO function and action the change.

# Function to set the time zone information for user.
function Set-TimeZone {
    foreach ($group in $script:groups.GetEnumerator()) {
        $groupMembers = Get-MgGroupMemberAsUser -GroupId $group.Key | Select-Object -ExpandProperty UserPrincipalName
        foreach ($groupMember in $groupMembers) {
            if ($groupMember -eq $script:user) {
                Write-Host "Setting $($script:user) mailbox time zones to $($group.Value)" -ForegroundColor Cyan
                Set-MailboxRegionalConfiguration -Identity $script:user -TimeZone "$($group.Value)"
                Set-MailboxCalendarConfiguration -Identity $script:user -WorkingHoursTimeZone "$($group.Value)"
            } 
        }
    }
}

# Function to check user's mailbox timezone settings.
function Get-UserTimeZones {
    foreach ($group in $script:groups.GetEnumerator()) {
        $groupMembers = Get-MgGroupMemberAsUser -GroupId $group.Key | Select-Object -ExpandProperty UserPrincipalName
        foreach ($groupMember in $groupMembers) {
            $script:user = $groupMember
            $mailboxRegionalConfig = Get-MailboxRegionalConfiguration -Identity $script:user | Select-Object -ExpandProperty TimeZone
            $mailboxCalendarConfig = Get-MailboxCalendarConfiguration -Identity $script:user | Select-Object -ExpandProperty WorkingHoursTimeZone

            # Checking if regional configuration matches calendar configuration.
            if ($mailboxRegionalConfig -ne $mailboxCalendarConfig) {
                Write-Host "[!] Regional Configuration ($mailboxRegionalConfig) and Calendar Configuration ($mailboxCalendarConfig) for $script:user are mismatched..." -ForegroundColor Red
                Set-TimeZone
            # Checking if user's regional and calendar config match security group membership
            } elseif (($mailboxRegionalConfig -ne $group.Value) -and ($mailboxCalendarConfig -ne $group.Value)) {
                Write-Host "[!] User's time zone does not match security group membership.. setting user to proper time zone." -ForegroundColor Red
                Set-TimeZone
            }
        }
    }
}

Get-UserTimeZones
