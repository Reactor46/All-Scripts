# api: multitool
# version: 0.1
# title: mini tools
# description: Small PS cmdlet collection
# type: inline
# category: info
# hidden: 0
# param: machine,username,minicmdlets
# config: {}
# 
# Various functions


Param(
    $machine = (Read-Host "Machine"),
    $user = (Read-Host "User"),
    $cmd = (Read-Host "minicmdlets")
)

#-- run
Write-Host -f Green "❏ $cmd"
$cmd = $cmd -replace '\$(machine|host|computer)',"'$machine'"
$cmd = $cmd -replace '\$(username|user|account)',"'$user'"
(Invoke-Expression $cmd | Out-String).trim() | Write-Host
