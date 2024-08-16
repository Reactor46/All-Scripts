http://technet.microsoft.com/en-us/library/bb310789.aspx

Get-RoleGroupMember -Identity "Public Folder Management" 
Add-RoleGroupMember -Identity "Public Folder Management" -Member Tony

Get-PublicFolderClientPermission -Identity "\Marketing" | fl
Get-PublicFolderClientPermission -Identity "\Marketing\EastCoast" -User David

Add-PublicFolderClientPermission -Identity "\Marketing\West Coast" -AccessRights PublishingEditor -User Kim