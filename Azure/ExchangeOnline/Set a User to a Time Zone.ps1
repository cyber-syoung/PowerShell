# Script settings
Set-ExecutionPolicy Bypass CurrentUser
$WarningPreference = 'SilentlyContinue'

# Welcome!
# Created by Sam Young.
Write-Host "`n###########################################################" -ForegroundColor Cyan
Write-Host "#                                                         #" -ForegroundColor Cyan
Write-Host "#           Set an individual user's time zone            #" -ForegroundColor Cyan
Write-Host "#                  Version 1, 8/14/2024                   #" -ForegroundColor Cyan
Write-Host "#                                                         #" -ForegroundColor Cyan
Write-Host "###########################################################`n" -ForegroundColor Cyan
Start-Sleep 1


# Installing Microsoft Graph Authentication module
Write-Host "Confirming installation of Microsoft Graph.. this may take a second." -ForegroundColor Blue
try {
    Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Write-Host "Microsoft Graph Authentication module installed and imported successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to install or import Microsoft Graph Authentication module. $_" -ForegroundColor Red
    Start-Sleep 3
    exit
}


# Connecting to Microsoft Graph
Write-Host "`nConnecting to graph.." -ForegroundColor Blue
try {
    Import-Module Microsoft.Graph.Authentication
    Connect-MgGraph -Scopes User.Read.All -NoWelcome -ErrorAction Stop
    Write-Host "Connected to Graph!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Microsoft Graph. $_" -ForegroundColor Red
    Start-Sleep 
    exit
}


# Installing Exchange Online Management module
Write-Host "`nConfirming installation of Microsoft Exchange Online.. this may take a second." -ForegroundColor Blue
try {
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "Exchange Online Management module installed and imported successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to install or import Exchange Online Management module. $_" -ForegroundColor Red
    Start-Sleep 3
    exit
}


# Connecting to Exchange Online
Write-Host "`nConnecting to exchange.." -ForegroundColor Blue
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected to Exchange Online!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Exchange Online. $_" -ForegroundColor Red
    Start-Sleep 3
    exit
}

# Part 4, checking to see if user wants to exit
function Exit-Script {
    $exitConfirmation = Read-Host "Would you like to configure another user's time zone? (y/n)"

    switch ($exitConfirmation) {
        "y" {
            Start-Sleep 1  
            Get-UserPrompt
        }
        "n" {
            Write-Host "`nGoodbye!" -ForegroundColor Yellow
            Start-Sleep 1
            exit
        }
        Default {
            Write-Host "Unknown response, please try again!" -ForegroundColor Red
            Exit-Script
        }
    }
}

# Part 3, takes care of actually setting the user to a time zone.
function Set-UserTimeZone {
    $timezones = @{
        "EST" = "Eastern Standard Time"
        "CST" = "Central Standard Time"
        "MST" = "Mountain Standard Time"
        "PST" = "Pacific Standard Time"
    }

    $tzConfirm = Read-Host "You have chosen $script:tzPrompt ($($timezones[$script:tzPrompt])), is this correct? (y/n)"
    switch ($tzConfirm) {
        "y" {  
            $timezone = $($timezones[$script:tzPrompt])

            try {
                Set-MailboxCalendarConfiguration -Identity $script:user -WorkingHoursTimeZone $timezone -ErrorAction Stop
                Write-Host "Setting Calendar Configuration for $script:user to $timezone..." -ForegroundColor Yellow
            } catch {
                Write-Host "Failed to set Calendar Configuration. $_" -ForegroundColor Red
            }

            try {
                Set-MailboxRegionalConfiguration -Identity $script:user -TimeZone $timezone -ErrorAction Stop
                Write-Host "Setting Regional Configuration for $script:user to $timezone...`n" -ForegroundColor Yellow
            } catch {
                Write-Host "Failed to set Regional Configuration. $_" -ForegroundColor Red
            }

            $calendar = Get-MailboxCalendarConfiguration -Identity $script:user | Select-Object -ExpandProperty WorkingHoursTimeZone
            $regional = Get-MailboxRegionalConfiguration -Identity $script:user | Select-Object -ExpandProperty TimeZone
            
            if (($calendar -ne $timezone) -or ($regional -ne $timezone)) {
                Write-Host "`nFailed to set time zone.. Trying again." -ForegroundColor Red
                Set-UserTimeZone
            } elseif (($calendar -ne $timezone) -and ($regional -ne $timezone)) {
                Write-Host "`nFailed to set time zone.. Trying again." -ForegroundColor Red
                Set-UserTimeZone
            } else {
                Write-Host "User's mailbox time zones have been successfully set to:`nCalendar: $calendar.`nRegional: $regional.`n" -ForegroundColor Green
                Exit-Script
            }
        }
        "n" {
            Get-TimeZone
        }
        Default {
            Write-Host "Unknown input, try again."
            Set-UserTimeZone
        }
    }


    $timezone = $($timezones[$script:tzPrompt])

    Set-MailboxCalendarConfiguration -Identity $script:user -WorkingHoursTimeZone $timezone 
    Write-Host "Calendar Configuration for $script:user has been set to $timezone." -ForegroundColor Green
    Set-MailboxRegionalConfiguration -Identity $script:user -TimeZone $timezone 
    Write-Host "Regional Configuration for $script:user has been set to $timezone.`n" -ForegroundColor Green

    Exit-Script
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
        try {
            #-ErrorAction Stop 
            $userConfirm = Get-MgUser -UserId $script:userPrompt | Select-Object -ExpandProperty Mail
            if ($null -ne $userConfirm) {
                $script:user = $script:userPrompt
                Write-Host "Found user: $script:user" -ForegroundColor Green
                $repeat = $false
                Get-TimeZone
            } else {
                Write-Host "ERROR: Unknown user, please try again." -ForegroundColor Red
            }
        } catch {
            Write-Host "Error retrieving user. $_" -ForegroundColor Red
        }
    }
}

Get-UserPrompt
