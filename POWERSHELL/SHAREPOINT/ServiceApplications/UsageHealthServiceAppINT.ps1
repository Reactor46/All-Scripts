################################################################################################# 
## Kelsey-Seybold SharePoint Server 2010 Central Admin Application Server Services Installation## 
################################################################################################# 
Set-ExecutionPolicy Unrestricted
  
Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue 
  
## Settings you may want to change depending on the install## 
$databaseServerName = "SPINT\Internal"

## Service Application Service Names ## 
$usageSAName = "Usage and Health Data Collection Service"

###########################################################################################################################################  
Write-Host "Creating Usage Service and Proxy..."
# No application pool creation required for this service application.
$serviceInstance = Get-SPUsageService
New-SPUsageApplication -Name $usageSAName -DatabaseServer $databaseServerName -DatabaseName "SP2013_INT_KSC_Intranet_UsageDB" -UsageService $serviceInstance > $null
###########################################################################################################################################

Remove-PSSnapin Microsoft.SharePoint.Powershell