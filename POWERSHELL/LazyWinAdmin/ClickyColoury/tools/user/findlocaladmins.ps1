# api: multitool
# version: 0.5
# title: Detect admins
# description: list local admins on remote machine
# type: inline
# category: user
# img: users.png
# hidden: 0
# key: i13|localadmins
# config: {}
# 
# list local admins on computer
# - via WMI win32_groupuser
# - shortened to DOMAIN\UserName


Param($machine = (Read-Host "Machine"))

Write-Host -f Green "WMI query 'win32_groupuser'..."
$admins = Gwmi win32_groupuser –computer $machine

Write-Host -f Green "Extract account infos"
$admins = $admins |? {$_.groupcomponent –like '*"Administrators"'}
$admins | ? {
    $_.PartComponent -match 'Domain="(.*)",Name="(.*)"'
} | % { 
    Write-Host -f Yellow "♞ $($matches[1])\$($matches[2])"
}

