Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Web application URL and CSV File location Variables
$WebAppURL="https://wiki.kscpulse.com"
$CSVFile="D:\Tempfiles\wikiSitesList.csv"
 
#Get list of site collections in a web application powershell
Get-SPWebApplication $WebAppURL | Get-SPSite -Limit All | ForEach-Object {
    New-Object -TypeName PSObject -Property @{
             SiteName = $_.RootWeb.Title
             Url = $_.Url
             DatabaseName = $_.ContentDatabase.Name }
} | Export-CSV $CSVFile -NoTypeInformation


#Read more: https://www.sharepointdiary.com/2016/10/get-all-site-collections-in-web-application-using-powershell.html#ixzz85xKAjhUR