# api: multitool
# version: 1.0
# title: [WMI] $user
# description: Show SID and other AD user properties
# type: inline
# category: wmi
# hidden: 0
# key: w1|sid|wmi_user|userid
# config: {}
# 
# Get detailed user info (such as SID for AD name) via WMI query.


Param($user = (Read-Host User))

[wmi] "win32_userAccount.Domain='$($cfg.domain)',Name='$user'" | Format-List -Prop *

