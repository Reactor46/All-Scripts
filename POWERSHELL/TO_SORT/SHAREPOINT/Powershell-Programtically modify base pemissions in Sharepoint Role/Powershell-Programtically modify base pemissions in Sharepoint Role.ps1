# SharePoint DLL
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
if ( (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

$wapp = $args[0]
$wsite = $args[1]

$webApplicationURL = $wapp

$webApp = Get-SPWebApplication $webApplicationURL

if($webApp -ne $null)
{
	write-host "Web Application : " $webApp  

	foreach($siteColl in $webApp.Sites)
	{
		if($siteColl -ne $null)
		{
			write-host "Site Collection : " $siteColl  
			foreach($subWeb in $siteColl.AllWebs)
			{
				if($subWeb -ne $null)
				{
				   
				   
					   if($subWeb.Url -like "$wsite*")
					   {
						if($subWeb.HasUniqueRoleAssignments -eq $True -and $subWeb.HasUniqueRoleDefinitions -eq $True)
						{
						    Write-Host "WebSite has unique permission/role : " $subWeb  " Url : "  $subWeb.Url 
							$roleColl=$subWeb.RoleDefinitions
								foreach($role in $roleColl)
								{
									if($role.name -ne $null)
									{

										write-host "--- Ediatble Role --- "  $role.Name
													
										$ExistingPermission = $role.BasePermissions

										$role.BasePermissions = "$ExistingPermission" , "UseRemoteAPIs, UseClientIntegration,  ViewFormPages"
										$role.Update()
									
									}
									else
									{
										write-host "XXX Not editable role XXX" $role.Name
										continue;
									}
								}
						
						}
						else
						{
							Write-Host "WebSite has inherited permission : "$subWeb  " Url : "  $subWeb.Url 
						}
					} 
					
				   $subWeb.Dispose()
				}
				else
				{
				   Echo $subWeb "does not exist"
				}
			}
			$siteColl.Dispose()
		}
		else
		{
		 Echo $siteColl "does not exist"
		}
	}
}
else
{
Echo $webApplicationURL "does not exist, check the WebApplication name"
}

Remove-PsSnapin Microsoft.SharePoint.PowerShell

Echo Finish
