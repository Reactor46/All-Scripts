# api: multitool
# version: 0.4
# title: Network adapters
# description: Scans remote PC via WMI win32_networkadapters
# type: inline
# category: wmi
# tag: proxy
# hidden: 0
# key: w4|network|network-?adapters|ada?pt[ers]*|nwa
# config: {}
# 
# Useful for detecting parallel LAN and WLAN connections.
#
#  → proxy issues
#


Param($machine = (Read-Host "Machine"));

$adapters = Get-WMIObject win32_networkadapter -filter "netconnectionstatus=2" -ComputerName $machine
Format-List -InputObject $adapters -Property NetConnectionID,Name,MACaddress,ServiceName,InterfaceIndex
