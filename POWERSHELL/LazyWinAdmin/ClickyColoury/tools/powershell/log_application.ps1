# api: multitool
# version: 0.1
# title: EventLog Application
# description: Get-EventLog
# type: inline
# category: powershell
# icon: events
# param: machine,logname
# key: i22|eventlog
# hidden: 0
# 
# Get-EventLog -computer $machine
#  ❏ prints -Newest 20 entres
#  ❏ use [LogName] combobox to switch to Security or System logs

Param(
    $machine = (Read-Host "Computer"),
    $logname = (Read-Host "LogName")
)

Get-EventLog -Computer $machine -Newest 20 -LogName $logname | FT Index,Time,Message -Auto -Wrap

