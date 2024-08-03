# The purpose of this script was to grab all of the users found with PST set as their timezone in mailbox configuration, and grab their object creation date. 
# Connect-MgGraph
# Connect-ExchangeOnline

$cred = Get-Credential
$server = (nslookup $env:USERDNSDOMAIN)[0].split(' ')[2]

$script:pstReport = @{}

function Start-Main {
    try {
        $userCreate = Get-ADUser -Filter "UserPrincipalName -eq '$script:upn'" -Server $server -Credential $cred -Properties * | Select-Object -ExpandProperty whenCreated
        $script:pstReport[$script:upn] = "$userCreate"
    }
    catch {
        if ($_.ErrorDetails.Message -like "*The specific mailbox Identity:*") {
            continue
        }
        Write-Host "ERROR: $_" -ForegroundColor Red
    }
}

# Define import file
$script:csvFilePath = "C:\Script\temp\Time Zone Project\PST Users.csv"

# Define export path for CSV
$script:csvExportPath = "C:\Script\temp\Time Zone Project\PST User Creation.csv"

# Iterate through csv file
Import-Csv -Path $csvFilePath | ForEach-Object {
    $script:upn = $_.UserPrincipalName
    Write-Host "`nGetting creation date for $script:upn"
    Start-Main
}

# Export results
$script:pstReport.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        UPN = $_.Key
        whenCreate = $_.Value
    }
} | Export-Csv -Path $csvExportPath -NoTypeInformation
