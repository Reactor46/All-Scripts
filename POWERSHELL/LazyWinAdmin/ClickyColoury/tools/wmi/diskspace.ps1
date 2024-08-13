# api: multitool
# version: 0.1
# title: Disk space
# description: List logicaldisks / free space
# type: inline
# category: wmi
# hidden: 0
# key: w3|service|services
# config: {}
# 
# list services on remote machine


Param($machine = (Read-Host "Machine"))

Get-WmiObject win32_logicaldisk -computer $machine | FL


