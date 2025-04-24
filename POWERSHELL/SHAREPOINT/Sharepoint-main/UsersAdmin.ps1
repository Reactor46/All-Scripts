Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Add farm admiN
$UserID="AD\NPF1A0005b8SPSha"
 
#*** Add User to SharePoint 2010 Farm Administrator Group ***
#Get Central Admin Web App
$CAWebApp = Get-SPWebApplication -IncludeCentralAdministration | where-object {$_.DisplayName -eq "SharePoint Central Administration v4"} 
#Get Central Admin site
$CAWeb = Get-SPweb($CAWebApp.Url) 
#Get Farm Administrators Group
$FarmAdminGroup = $CAWeb.SiteGroups["Farm Administrators"] 
#Add user to the Group
$FarmAdminGroup.AddUser($UserID,"",$UserID , "")
Write-Host "User: $($UserID) has been added to Farm Administrators Group!"
$CAWeb.Dispose()
 
#***Add user to Web App Policy ***
   Get-SPWebApplication | foreach-object {
                $WebAppPolicy = $_.Policies.Add($UserID, $UserID)
                $PolicyRole = $_.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl)
                $WebAppPolicy.PolicyRoleBindings.Add($PolicyRole)
                $_.Update()
    Write-Host "Added user to $($_.URL)"
                } 
 
#*** Grant Shell Admin Access *** 
#Get All SharePoint Databases and Add user into Shell Admin access
Get-SPDatabase | Add-SPShellAdmin -Username $UserID

## Find Farm Administrators using PowerShell

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Get Central Admin Web App
$CAWebApp = Get-SPWebApplication -IncludeCentralAdministration | where-object {$_.DisplayName -eq "SharePoint Central Administration v4"} 
 
#Get Central Admin site
$CAWeb = Get-SPweb($CAWebApp.Url) 
 
$FarmAdminGroup = $CAWeb.SiteGroups["Farm Administrators"] 
foreach ($Admin in $FarmAdminGroup.users) 
{ 
   write-host $Admin.LoginName
}     