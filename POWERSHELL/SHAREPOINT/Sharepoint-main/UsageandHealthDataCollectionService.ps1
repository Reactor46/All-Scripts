Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Configuration Variables
$UsageSAName = "Hosting Farm Usage and Health Data Collection Service"
$DatabaseName = "HostingFarm_Usage_SA"
$DatabaseServer="agl-sppre-d.ad.ing.net"
$UsageLogPath = "D:\Log\Sharepoint"
 
#Get the Usage Service Application Service Instance
$ServiceInstance = Get-SPUsageService
 
#Create new Usage Service application
New-SPUsageApplication -Name $UsageSAName -DatabaseServer $DatabaseServer -DatabaseName $DatabaseName -UsageService $ServiceInstance #> $null
 
#Create Service Application Proxy..." 
$proxy = Get-SPServiceApplicationProxy | where {$_.TypeName -eq "Usage and Health Data Collection Proxy"}
$proxy.Provision()
 
#Set Usage Service Application Option
Set-SPUsageService -LoggingEnabled 1 -UsageLogLocation $UsageLogPath -UsageLogMaxSpaceGB 1 
