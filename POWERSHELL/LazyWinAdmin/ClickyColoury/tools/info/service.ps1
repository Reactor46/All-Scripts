# api: multitool
# version: 0.2
# title: Services
# description: list services on remote machine
# type: inline
# category: info
# hidden: 0
# key: i8|service|services
# config: {}
# 
# list services on remote machine


Param($machine = (Read-Host "Machine"))

Get-Service -computer $machine | Select-Object status,name,description | FT


