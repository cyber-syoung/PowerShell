write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host -ForegroundColor Blue "Visual Studio Code Version"
write-host "Computer Information Collector; the Third Version"
write-host "Written by Charlie Collins and Sam Young"
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host -ForegroundColor red "If you inquire about a TV, expect to change some fields."
write-host "Date of current update written: 6/8/2023"
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""
Start-Sleep 2

$NoExitEver= 1
while($NoExitEver -eq 1){

write-host "Select desired computer to collect. When finish, press 'Ctrl' and 'C' at same time."
Start-Sleep 1
write-host ""
$NameHere= read-host "Select name"
write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""
write-host "Checking for computer..."
start-sleep 1
write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""

$Ping?= Test-Connection -ComputerName $NameHere -Count 2 -Quiet

if($Ping? -eq $True){
New-Item -Path C:\ -name "ComputerInfoReport" -ItemType Directory -ErrorAction SilentlyContinue

Invoke-Command -ComputerName $NameHere -ScriptBlock {Get-ComputerInfo | Select-object csname, csmodel, csusername, biosseralnumber} > C:\ComputerInfoReport\$($NameHere).txt

write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""
write-host "Found computer..."
start-sleep 1
write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""
write-host "Gathering information..."
start-sleep 1
write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""

# Take the information from the txt file.
$computerModel = Get-Item -Path C:\ComputerInfoReport\$($NameHere).txt | Select-String -Pattern "CsModel" | Out-String # "C:\ComputerInfoReport\tvday-it101.txt:3:CsModel         :"
$computerSerial = Get-Item -Path C:\ComputerInfoReport\$($NameHere).txt | Select-String -Pattern "BiosSeralNumber" | Out-String # "C:\ComputerInfoReport\tvday-it101.txt:5:BiosSeralNumber :"
$computerHost = Get-Item -Path C:\ComputerInfoReport\$($NameHere).txt | Select-String -Pattern "CsUserName" | Out-String # "C:\ComputerInfoReport\tvday-it101.txt:4:CsUserName      : SCPC\"

$modelText = "C:\ComputerInfoReport\$($NameHere).txt:3:CsModel         : "
$serialText = "C:\ComputerInfoReport\$($NameHere).txt:5:BiosSeralNumber : "
$hostText = "C:\ComputerInfoReport\$($NameHere).txt:4:CsUserName      : SCPC\"

write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""
write-host "Gathering serial number..."
start-sleep 1
write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""

# Fix formatting for serial number
$serial = $computerSerial.replace($serialText,"")

write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""
write-host "Gathering username..."
start-sleep 1
write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""

# Fix formatting for username
$userReplace1 = $computerHost.replace($hostText,"")
$userReplace2 = $userReplace1.trim()
$userName = Get-ADUser -Filter {sAMAccountName -eq $userReplace2} | Select-Object -ExpandProperty "Name"

# Fix formatting for computer model
$model = $computerModel.replace($modelText,"")

write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""
write-host "Gathering location..."
start-sleep 1
write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""

# Location
$officeMap = @{
    # DAY
    "Dayton Plant - 9" = "DAY"
    "Plumbers" = "DAY"
    "Corp Dayton Plant - 9" = "DAY"
    "Day" = "DAY"
    # CCI
    "CCI Plant - 16" = "CCI"
    "Corp CCI Plant - 16 " = "CCI"
    "16-CCI" = "CCI"
    "CCI" = "CCI"
    # WCH
    "WCH Plant - 1" = "WCH"
    "Corp WCH Plant - 1" = "WCH"
    "Corp Bundy" = "BUN"
    "Corp Visador" = "VIS"
    "Visador" = "VIS"
    "Bundy Plant - 12" = "BUN"
    "01 - WCH" = "WCH"
    "01 -WCH" = "WCH"
    # BWB
    "Corp BWB" = "BWB"
    "Brandworthy Plant- 18" = "BWB"
    "18-BWB" = "BWB"
    # Cincinnati
    "18-CIN" = "CIN"
    "Cincinnati Plant - 10" = "CIN"
    # Corporate
    "05 - CORP" = "GTI"
    "CORP GTI" = "GTI"
    "GTI" = "GTI"
    # KAN
    "Corp Kansas Plant - 4" = "KAN"
    "Kansas Plant - 4" = "KAN"
    "04 - KAN" = "KAN"
    "04 - KAN FLOOR PC" = "KAN"
    "11 - CAR" = "CAR"
    "Carthage Plant - 11" = "CAR"
    "CAR" = "CAR"
    # Wingate
    "Wingate" = "WIN"
    "Wingate North" = "WIN"
    "05c - CORP" = "WIN"
}

# Get user location
$userOffice = Get-ADUser -Identity "$userReplace2" -Properties * | Select-Object -ExpandProperty "physicalDeliveryOfficeName"
if ($officeMap.Contains(($userOffice))) {
    $userLocation = $officeMap[$userOffice]
} elseif ($userOffice = "","NA") {
    Write-Host "$userReplace2 does not have a valid office location. Please review."
}
else {
    # If the user's office is not found in the hashtable
    Write-Host "Error: No matching office found for $($userName)"
}
Write-Host -ForegroundColor green "Computer information: $($serial.trim()) - $userName - $userLocation - $($model.trim())"
write-host ""
write-host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
write-host ""}
elseif ($Ping? -eq $False) 
{write-host -ForegroundColor red "ERROR: PC may be down or name was typed wrong. Please check your spelling."} 
write-host ""
}
