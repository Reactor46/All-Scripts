# api: multitool
# version: 0.3
# title: get RDP permissions
# description: on remote PC
# type: inline
# depends: funcs_base
# category: info
# img: registry
# hidden: 0
# 
# SYSTEM\CurrentControlSet\\Control\\Terminal Server\fDenyTSConnections


$v = Get-RemoteRegistry("\\$machine\HKLM\SYSTEM\CurrentControlSet\\Control\\Terminal Server\fDenyTSConnections")
Write-Host "RDP = $(!$v)"


