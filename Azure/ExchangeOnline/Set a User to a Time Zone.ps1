# Connect to Microsoft for some magic
Write-Host "Connecting to exchange.." -ForegroundColor Blue
Connect-ExchangeOnline -ShowBanner:$false 
Write-Host "Connected to exchange!" -ForegroundColor Green
Write-Host "`nConnecting to graph.." -ForegroundColor Blue
Connect-MgGraph -NoWelcome
Write-Host "Connected to Graph!" -ForegroundColor Green

# Part 3, takes care of actually setting the user to a time zone.
function Set-UserTimeZone {
    $timezones = @{
        "EST" = "Eastern Standard Time"
        "CST" = "Central Standard Time"
        "MST" = "Mountain Standard Time"
        "PST" = "Pacific Standard Time"
    }

    $timezone = $($timezones[$script:tzPrompt])

    Set-MailboxCalendarConfiguration -Identity $script:user -WorkingHoursTimeZone $timezone 
    Write-Host "Calendar Configuration for $script:user has been set to $timezone." -ForegroundColor Green
    Set-MailboxRegionalConfiguration -Identity $script:user -TimeZone $timezone 
    Write-Host "Regional Configuration for $script:user has been set to $timezone.`n" -ForegroundColor Green
}

# Part 2, asks the technician what time zone would they like to set the user to. 
function Get-TimeZone {
    $repeat = $true
    while ($repeat) {
        $script:tzPrompt = Read-Host "`nWhat is the Time Zone? [I.e. EST]"

        switch ($script:tzPrompt) {
            "EST" {
                Set-UserTimeZone
                $repeat = $false
            }
            "CST" {
                Set-UserTimeZone
                $repeat = $false
            }
            "MST" {
                Set-UserTimeZone
                $repeat = $false
            }
            "PST" {
                Set-UserTimeZone
                $repeat = $false
            }
            Default {
                Write-Host "ERROR: Unknown time zone, please try again." -ForegroundColor Red
            }
        }
    }
}

# Part 1, requests the technician to provide the user's UPN to make changes to.
function Get-UserPrompt {
    $repeat = $true
    while ($repeat) {
        $script:userPrompt = Read-Host "`nWho is the user? [UPN]"
        $userConfirm = Get-MgUser -UserId $script:userPrompt | Select-Object -ExpandProperty Mail
        if ($null -ne $userConfirm) {
            $script:user = $script:userPrompt
            Write-Host "Found user: $script:user" -ForegroundColor Green
            $repeat = $false
            Get-TimeZone
        } else {
            Write-Host "ERROR: Unknown user, please try again." -ForegroundColor Red
        }
    }
}

Get-UserPrompt



