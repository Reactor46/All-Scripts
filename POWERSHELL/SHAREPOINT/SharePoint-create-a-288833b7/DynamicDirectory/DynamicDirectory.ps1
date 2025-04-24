Add-PSSnapin Microsoft.SharePoint.Powershell -ea 0

# settings
$webApplication = "http://lrd02sha:20000/"
$excludeUrl = @("http://lrd02sha:20000")

$targetDirectory = "http://lrd02sha:20000"
$basePermissionLevel = "Read Only Items"
$linksList = "Workspaces"
# end settings

# this function resets the unique role assignments of an item and breaks the role inheritance
# afterwards we are going to add a single entry for every user with permissions on the target web assigned to the link item
function AssignPermissions($web, $listItem) {
    if ($listItem.HasUniqueRoleAssignments) {
      $listItem.ResetRoleInheritance()
    }
    $listItem.BreakRoleInheritance($false)
    
    Get-SPUser -Web $web.Url | where { $_.Roles.Count -gt 0 -or $_.Groups.Count -gt 0 } | ForEach-Object {
        $roleDefinition = $rootWeb.RoleDefinitions[$basePermissionLevel]
        $roleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($_.LoginName, $_.Email, $_.Name, $_.Notes)
        $roleAssignment.RoleDefinitionBindings.Add($roleDefinition)
        $listItem.RoleAssignments.Add($roleAssignment)
    }
    
    $listItem.Update()
}

# get the root web and list
$rootWeb = Get-SPWeb $targetDirectory
$list = $rootWeb.Lists[$linksList]

# check if the Permissionssetting is available, if not, then create one
if($rootWeb.RoleDefinitions[$basePermissionLevel] -eq $null)
{
    # Role Definition named "Add Only" does not yet exist
    $spRoleDefinition = New-Object Microsoft.SharePoint.SPRoleDefinition
    $spRoleDefinition.Name = $basePermissionLevel
    $spRoleDefinition.BasePermissions = "ViewListItems, Open, ViewPages"
    $rootWeb.RoleDefinitions.Add($spRoleDefinition)
}

# loop through all spsites in the specified web application and exclude the urls located under $excludeUrl
Get-SPWebApplication $webApplication | Get-SPSite -Limit All | where { $excludeUrl -notcontains $_.Url } | Foreach-Object  {
    # get the user of the web and verify in the links list of the directory if there is already an entry
    $myWeb = Get-SPWeb $_.Url
    $query=new-object Microsoft.SharePoint.SPQuery
    $caml = '<Where>
                <BeginsWith>
                    <FieldRef Name="URL"/>
                        <Value Type="URL">{0}</Value>
                </BeginsWith>
             </Where>' -f $myWeb.ServerRelativeUrl

    $query.Query=$caml
    $items = $list.GetItems($query);
    
    # if one or more items are returned, modify them accordingly. else create a new entry in the directory
    if ($items.Count -gt 0) {
        $items | ForEach-Object {
            $listItem = $_
            
            # modify the exiting item url by specifying the url and title of the web
            $listItem["URL"] = '{0}, {1}' -f $myWeb.Url, $myWeb.Title
            $listItem.Update()
            
            # re-assign the permissions to this item for security trimming
            AssignPermissions $myWeb $listItem
        }
    }
    else {
        # create a new item and assign the url and title of the current web
        $listItem = $list.Items.Add()
        $listItem["URL"] = '{0}, {1}' -f $myWeb.Url, $myWeb.Title
        $listItem.Update()
        
        # assign the permissions to this item for security trimming
        AssignPermissions $myWeb $listItem
    }
}