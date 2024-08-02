# Connect-MgGraph

# Purpose of this function is to remove the license from the UPN from the CSV at the input path address.
function Remove-DirectLicenses {
    # Grabbing the userID from UPN.
    $userID = Get-MgUser -UserId $script:upn -Property * | Select-Object -ExpandProperty Id -ErrorAction Stop
    try {
        # Grabbing the license SKU ID.
        $lic = Get-MgSubscribedSku -All -ErrorAction Stop | Where-Object SkuPartNumber -eq "$script:license" | Select-Object -ExpandProperty SkuId -ErrorAction Stop
        # Removing the license from the user.
        Set-MgUserLicense -UserId $userID -RemoveLicenses $lic -AddLicenses @{} -ErrorAction Stop
    }
    catch {
        # Cleaning up to prevent error spamming.
        if ($_.Exception.Message -like "*User license is inherited from a group membership*") {
            Write-Host "User is receiving license from SG... skipping." -ForegroundColor Yellow
        } else {
            Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Manually set input path.
$csvFilePath = "C:\Script\temp\License Reports\Power Automate Free.csv"

# Importing the CSV from the input path.
Import-Csv -Path $csvFilePath | ForEach-Object {
    $script:upn = $_.UserPrincipalName
    $script:license = $_.License
    Write-Host "Removing $script:license from $($script:upn)'s account." -ForegroundColor Cyan
    Remove-DirectLicenses
}
