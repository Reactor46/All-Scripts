<# 
.Script Information
    Creator:  Cash Conway
    Date Created:  11/10/2014
    Purpose:  Script to shutdown services on SAS services

    Run under Advisor Credentials

    Script will take around 10 minutes to run, built in time delay on shutdown to ensure all services get stopped.
    
 
#>



Write-Host "SHUTDOWN SAS services"
Write-Host " "
Write-Host " "

$continue = Read-Host "Do you really want to STOP the SAS Services (Y/N)?"
Write-Host " "
Write-Host " "

If ($continue -eq "Y")  {
        write-host 1 - Shutdown SAS Deployment Agent on lassasmt01
        Write-host
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS Deployment Agent' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 2 - Shutdown SAS Deployment Agent on lassasc01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS Deployment Agent' -ComputerName LASSASC01  | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 3 - Shutdown SAS [ConfigMid-Lev1] SAS Environment Manager Agent on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 4- Shutdown SAS [Config-Lev1] SAS Environment Manager Agent on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 5 - Shutdown SAS [Config-Lev1] SAS Environment Manager Agent on lassasc01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent' -ComputerName LASSASC01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 6 - Shutdown SAS [ConfigMid-Lev1] SAS Environment Manager on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 7 - Shutdown SAS [ConfigMid-Lev1] WebAppServer SASServer1_1 on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 8 - Shutdown SAS[ConfigMid-Lev1]httpd-WebServer on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 9 - Shutdown SAS [ConfigMid-Lev1] Cache Locator on port 41415 on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 10 - Shutdown SAS [Config-Lev1] Cache Locator on port 41415 on lassasc01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415' -ComputerName LASSASC01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 11 - Shutdown SAS [ConfigMid-Lev1] JMS Broker on port 61616 on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 12 - Shutdown SAS[Config-Lev1]DIPJobRunner on lassasc01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS[[]Config-Lev1[]]DIPJobRunner' -ComputerName LASSASC01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 13 - Shutdown SAS [Config-Lev1] Connect Spawner on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]Config-Lev1[]] Connect Spawner' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 14 - Shutdown SAS [Config-Lev1] Connect Spawner on lassasc01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]Config-Lev1[]] Connect Spawner' -ComputerName LASSASC01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 15 - Shutdown SAS [Config-Lev1] Object Spawner on lassasc01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]Config-Lev1[]] Object Spawner' -ComputerName LASSASC01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 16 - Shutdown SAS [Config-Lev1] SASMeta - Metadata Server on lassasmt01
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server' -ComputerName LASSASMT01 | Set-Service -Status Stopped
        Start-Sleep -s 60
      Write-host
        write-host 17 - Shutdown SAS [Config-Lev1] Web Infrastructure Platform Data Server
        write-host "**************************************************************************************"
      Write-host
        Get-Service -Name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server' -ComputerName LASSASC01 | Set-Service -Status Stopped
        Start-Sleep -s 60
        
}

Write-Host Verify all services have been stopped
