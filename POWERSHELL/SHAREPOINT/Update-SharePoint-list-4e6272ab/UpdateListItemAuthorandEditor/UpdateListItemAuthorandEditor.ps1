Add-PSSnapin Microsoft.SharePoint.PowerShell -EA 0
#specify new user - login name
$newuser = "support\paulmather"
#specify the web URL
$site = get-SPWeb http://vm353/pwa/ProjectBICenter
#specify the list name
$list = $site.Lists["Test"]
$listitems = $list.Items
$user = get-SPuser -Web $site -Identity $newuser
#loop through each list item and update
foreach ($listitem in $listitems)
    {
        
        $listitem["Author"] = $user
        $listitem["Editor"] = $user       
        $listitem.Update()
        write-host $listitem["Name"] "has been updated. The author and editor has been set to $user"
        
    }
$site.Update()
$site.Dispose()