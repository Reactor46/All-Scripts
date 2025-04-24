################################################################################################# 
## Kelsey-Seybold SharePoint Server 2010 Central Admin Application Server Services Installation## 
################################################################################################# 
Set-ExecutionPolicy Unrestricted
  
Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue 
  
## Settings you may want to change depending on the install## 
$databaseServerName = "SPINT\Internal"

## Service Application Service Names ## 
$stateSAName = "State Service"

##########################################################################################################################################  
Write-Host "Creating State Service and Proxy..."
# No application pool created for this service application
New-SPStateServiceDatabase -Name "SP2013_INT_KSC_Intranet_StateServiceDB" -DatabaseServer $databaseServerName | New-SPStateServiceApplication -Name $stateSAName | New-SPStateServiceApplicationProxy -Name "$stateSAName Proxy" -DefaultProxyGroup > $null
##########################################################################################################################################

Remove-PSSnapin Microsoft.SharePoint.Powershell