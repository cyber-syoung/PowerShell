# Connect to Microsoft Graph
# Connect-MgGraph

# Retrieve all users
$script:users = Get-MgUser -All | Select-Object -ExpandProperty UserPrincipalName

# Initialize the license report.
$script:licenseReport = @{}

# Prompt to ask for license name, needs to be exact.
$licensePrompt = Read-Host "License name?"

# Define the output path
$script:csvExportPath = "C:\Script\temp\License Reports\$licensePrompt.csv"

try {
    # Iterating through all users from Get-MgUser.
    foreach ($user in $script:users) {
        # Grabbing the user's licenses
        $licenses = Get-MgUserLicenseDetail -UserId $user | Select-Object -ExpandProperty SkuPartNumber
        # Checking for the requested license from licensePrompt.
        foreach ($lic in $licenses) {
            if ($lic -eq "$licensePrompt") {
                Write-Host "Found $user has $lic" -ForegroundColor Green
                $script:licenseReport[$user] = "$lic"
            }
        }
    }

    <#
    # Debugging: Check the content of the license report
    Write-Host "License Report Content:" -ForegroundColor Cyan
    $script:licenseReport.GetEnumerator() | ForEach-Object {
        Write-Host "$($_.Key): $($_.Value)" -ForegroundColor Cyan
    }#>
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
}

# Export results to CSV
try {
    Write-Host "Exporting to CSV at $script:csvExportPath" -ForegroundColor Yellow
    $script:licenseReport.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{
            UserPrincipalName = $_.Key
            License = $_.Value
        }
    } | Export-Csv -Path $script:csvExportPath -NoTypeInformation -Force
    Write-Host "Export completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to export to CSV: $_" -ForegroundColor
}
