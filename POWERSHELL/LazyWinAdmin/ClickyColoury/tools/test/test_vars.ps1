# api: multitool
# title: test vars
# description: $GUI.vars and Get-ParamVarsCmd interpol
# version: 0.0
# param: machine, username, field1, field2
# category: test
# type: window
#
# testy test

param($mach=0, $user=0, $field1=0, $field2=0)

Write-Host "host=$mach"
Write-Host "user=$user"
Write-Host "f1=$field1"
Write-Host "f2=$field2"
Read-host "Wait"
