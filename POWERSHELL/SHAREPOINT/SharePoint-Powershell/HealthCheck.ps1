$centralAdminURL = "http://fbv-spapp-t03:9999/"
$listName = "Review Problems and Solutions"
$spSourceWeb = Get-SPWeb $centralAdminURL
$spSourceList = $spSourceWeb.Lists[$listName]
$spSourceItems = $spSourceList.GetItems() | where {$_['Severity'] -ne "4 - Success"}
$spSourceItems | ForEach-Object {
   Write-Host $_['Title']
   }