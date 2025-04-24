################################################################################################# 
## Kelsey-Seybold SharePoint Server 2013 State Service Installation                            ## 
################################################################################################# 
Set-ExecutionPolicy Unrestricted
  
Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue 
  
## Settings you may want to change depending on the install## 
$databaseServerName = "SPProd\KSExternal"

## Service Application Service Names ## 
$stateSAName = "SharePoint Server ASP.Net Session State Service"

##########################################################################################################################################  
Write-Host "Creating State Service and Proxy..."
# No application pool created for this service application
New-SPStateServiceDatabase -Name "SP2013Ext_ASPNetStateService_Service" -DatabaseServer $databaseServerName | New-SPStateServiceApplication -Name $stateSAName | New-SPStateServiceApplicationProxy -Name "$stateSAName Proxy" -DefaultProxyGroup > $null
##########################################################################################################################################

Remove-PSSnapin Microsoft.SharePoint.Powershell