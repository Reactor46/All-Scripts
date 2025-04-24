# Change Global Search Center URL with PowerShell
#Read more: https://www.sharepointdiary.com/2017/02/set-global-search-center-in-sharepoint.html
#Get the Search Service Application
$SSA = Get-SPEnterpriseSearchServiceApplication
 
#Set global search center
$SSA.SearchCenterUrl = "https://Search-Center-URL/pages"
 
$SSA.Update()


