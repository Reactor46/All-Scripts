<# 
.Script Information
    Creator:  Cash Conway
    Date Created:  11/10/2014
    Purpose:  Script to shutdown services on SAS services

    Run under Advisor Credentials

    Script will take around 10 minutes to run, built in time delay on shutdown to ensure all services get Running.
    
 
#> 


Write-Host "START SAS services"
Write-Host " "
Write-Host " "

$continue = Read-Host "Do you really want to START the SAS Services (Y/N)?"

If ($continue -eq "Y")  {
    Get-Service -Name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server' -ComputerName LASSASC01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]Config-Lev1[]] Object Spawner' -ComputerName LASSASC01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]Config-Lev1[]] Connect Spawner' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]Config-Lev1[]] Connect Spawner' -ComputerName LASSASC01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS[[]Config-Lev1[]]DIPJobRunner' -ComputerName LASSASC01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415' -ComputerName LASSASC01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent' -ComputerName LASSASC01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent' -ComputerName LASSASMT01 | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS Deployment Agent' -ComputerName LASSASC01  | Set-Service -Status Running
    Start-Sleep -s 30
    Get-Service -Name 'SAS Deployment Agent' -ComputerName LASSASMT01 | Set-Service -Status Running

}












