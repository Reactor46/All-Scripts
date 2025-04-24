function GetSPWebUsers($SiteCollectionURL)
{
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") > $null
    $site = new-object Microsoft.SharePoint.SPSite($SiteCollectionURL)
    $web = $site.openweb()
    $siteUsers = $web.SiteUsers

    foreach($user in $siteUsers)
    {		
	    Write-Host " ------------------------------------- "
	    Write-Host "Site Collection URL:", $SiteCollectionURL
	    if($user.IsSiteAdmin -eq $true)
	    {
		    Write-Host "ADMIN: ", $user.LoginName
	    }
	    else
	    {
		    Write-Host "USER: ", $user.LoginName
	    }
	    Write-Host " ------------------------------------- "
    }	
    $web.Dispose()
    $site.Dispose()
}


function GetSPAllSPUsers($SiteCollectionURL,$SPListName)
{
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") > $null
    $site = new-object Microsoft.SharePoint.SPSite($SiteCollectionURL)	
    $web = $site.openweb()
    $list = $web.Lists[$SPListName]
    $siteCollUsers = $web.SiteUsers
    
    foreach($user in $siteCollUsers)
	    {
		    Write-Host " ------------------------------------- "
		    Write-Host "Site Collection URL:", $SiteCollectionURL
            if($list.DoesUserHavePermissions([Microsoft.SharePoint.SPBasePermissions]::ViewListItems,$user) -eq $true)
                {
                    Write-Host "User : ", $user.LoginName
                    Write-Host "Assigned Permissions : ", $list.GetUserEffectivePermissions($user.LoginName)
                }			
            Write-Host " ------------------------------------- "		
	    }
	
	    $web.Dispose()
	    $site.Dispose()
 }
 
