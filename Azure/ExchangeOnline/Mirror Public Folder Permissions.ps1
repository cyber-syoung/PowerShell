# Connect to ExchangeOnline
Connect-ExchangeOnline

# Grab all membership and permissions from group to mirror permissions from
$mpFromPrompt = Read-Host "What is the file path of the group you want to mirror permissions from? [FORMAT: \Training Team ]"
$mpFrom = Get-PublicFolderClientPermission -Identity "$mpFromPrompt" | Select-Object FolderName,User,AccessRights

$group1 = @()
foreach ($mpF in $mpFrom) {
    $group1 += $mpF
}


# Grab all membership and permissions from group to mirror permissions to
