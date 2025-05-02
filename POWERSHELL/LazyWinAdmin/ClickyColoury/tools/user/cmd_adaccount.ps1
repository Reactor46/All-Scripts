# api: multitool
# version: 1.0
# title: AD account
# description: Get detailed AD account information
# type: psexec
# category: user
# img: users.png
# hidden: 0
# key: 7|ad-?account|account
# config: -
#
# Get all propertires via Get-ADUser
#
# Alternatives:
#  · [ADSISearcher]
#  · WMI Win32_user
#

Param($user = (Read-Host "Username"))

#-- ActiveDirectory query
Get-ADUser $user -Properties * | SELECT SamAccountName,Enabled,LockedOut,PasswordExpired,PasswordLastSet,AccountExpirationDate | Format-List

