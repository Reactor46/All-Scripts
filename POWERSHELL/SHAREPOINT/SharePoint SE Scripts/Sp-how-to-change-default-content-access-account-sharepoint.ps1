# Read more: https://www.sharepointdiary.com/2016/09/how-to-change-default-content-access-account-sharepoint-2016-search.html#ixzz8BFCtbSDr
# PowerShell to Change Search Crawl Account in SharePoint 2016:
Add-PSSnapin Microsoft.SharePoint.Powershell -ErrorAction SilentlyContinue
 
#Set Default content access account
$AccountID = "Crescent\SP16_Crawl"
$Password = Read-Host -AsSecureString 
 
#Get Search service application
$SearchInstance = Get-SPEnterpriseSearchServiceApplication
 
#Set default content access account for crawl
Set-SPEnterpriseSearchServiceApplication -Identity $SearchInstance -DefaultContentAccessAccountName $AccountID -DefaultContentAccessAccountPassword $Password