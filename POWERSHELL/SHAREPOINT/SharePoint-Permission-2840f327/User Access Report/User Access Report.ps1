Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

 
#Function to retrieve Permission data - Coded by Salaudeen Rajack: http://www.SharePointDiary.com
Function Get-PermissionInfo([String]$UserID, [Microsoft.SharePoint.SPSecurableObject]$Object)
{
 #Object Array to hold Permission data
    $PermissionDataCollection = @()
 
 #Determine the given Object type and Get URL of it
    switch($Object.GetType().FullName)
 {
  "Microsoft.SharePoint.SPWeb"  { $ObjectType = "Site" ; $ObjectURL = $Object.URL }
  "Microsoft.SharePoint.SPListItem"
  {
   if($Object.Folder -ne $null)
   {
     $ObjectType = "Folder" ; $ObjectURL = "$($Object.Web.Url)/$($Object.Url)"
   }
   else
   {
    $ObjectType = "List Item"; $ObjectURL = "$($Object.Web.Url)/$($Object.Url)"
   }
  }
  #Microsoft.SharePoint.SPList, Microsoft.SharePoint.SPDocumentLibrary, "Microsoft.SharePoint.SPPictureLibrary",etc
  default { $ObjectType = "List/Library"; $ObjectURL = "$($Object.ParentWeb.Url)/$($Object.RootFolder.URL)" }
 }
  
 #Get Permissions of the user on given object - Such as: Web, List, Folder, ListItem
 $UserPermissionInfo = $Object.GetUserEffectivePermissionInfo($UserID)
 #Iterate through each permission and get the details
 foreach($UserRoleAssignment in $UserPermissionInfo.RoleAssignments)
 {
  #Get all permission levels assigned to User account directly or via SharePOint Group
  $UserPermissions=@()
        foreach ($UserRoleDefinition in $UserRoleAssignment.RoleDefinitionBindings)
        {
   #Exclude "Limited Accesses"
   if($UserRoleDefinition.Name -ne "Limited Access")
   {
          $UserPermissions += $UserRoleDefinition.Name
   }
        }
  
  #Determine Permissions granted directly or through SharePoint Group
  if($UserPermissions)
  {
   if($UserRoleAssignment.Member -is [Microsoft.SharePoint.SPGroup])  
   {
     $PermissionType = "Member of SharePoint Group: " + $UserRoleAssignment.Member.Name    
   }
   else
   {
    $PermissionType = "Direct Permission"
   }
   $UserPermissions = $UserPermissions -join ";" 
  
   #Create an object to hold storage data
         $PermissionData = New-Object PSObject
         $PermissionData | Add-Member -type NoteProperty -name "Object" -value $ObjectType
   $PermissionData | Add-Member -type NoteProperty -name "Title" -value $Object.Title
         $PermissionData | Add-Member -type NoteProperty -name "URL" -value $ObjectURL 
   $PermissionData | Add-Member -type NoteProperty -name "Permission Type" -value $PermissionType
   $PermissionData | Add-Member -type NoteProperty -name "Permissions" -value $UserPermissions
   $PermissionDataCollection += $PermissionData
  }  
 }
 Return $PermissionDataCollection
}
 
#Function to Generate Permission Report
Function Generate-PermissionReport($UserID, $WebAppURL, $ReportPath)
{
    #Output Report location, delete the file, If already exist!
    if (Test-Path $ReportPath)
     {
        Remove-Item $ReportPath
     }
   
   #Write Output Report CSV File Headers
  "Object, Title, URL, Permission Type, Permissions" | out-file $ReportPath
 
 ###Check Whether the Search Users is a Farm Administrator ###
 Write-host "Scanning Farm Administrators..."
   #Get the SharePoint Central Administration site
   $AdminWebApp = Get-SPwebapplication -includecentraladministration | where {$_.IsAdministrationWebApplication}
    $AdminSite = Get-SPWeb $AdminWebApp.Url
    $AdminGroupName = $AdminSite.AssociatedOwnerGroup
    $FarmAdminGroup = $AdminSite.SiteGroups[$AdminGroupName]
  
 #enumerate in farm adminidtrators groups
    foreach ($user in $FarmAdminGroup.users)
    {
     if($user.LoginName.Endswith($UserID,1)) #1 to Ignore Case
     {
       "Farm, $($AdminSite.Title), $($AdminWebApp.URL), Farm Administrators Group, Farm Administrator" | Out-File $ReportPath -Append
     }    
    }
  
 ### Check Web Application User Policies ###
 Write-host "Scanning Web Application Policies..."
  $WebApp = Get-SPWebApplication $WebAppURL
  
  foreach ($Policy in $WebApp.Policies)
  {
      #Check if the search users is member of the group
     if($Policy.UserName.EndsWith($UserID,1))
       {
       #Write-Host $Policy.UserName
        $PolicyRoles=@()
       foreach($Role in $Policy.PolicyRoleBindings)
       {
        $PolicyRoles+= $Role.Name +";"
       }
   #Send Data to CSV File
      "Web Application, $($WebApp.Name), $($WebApp.URL), Web Application Policy, $($PolicyRoles)" | Out-File $ReportPath -Append
   }
  }
 
 #Convert UserID Into Claims format - If WebApp is claims based! Domain\User to i:0#.w|Domain\User
    if($WebApp.UseClaimsAuthentication)
    {
        $ClaimsUserID = (New-SPClaimsPrincipal -identity $UserID -identitytype 1).ToEncodedString()
    }
  
 #Get all Site collections of given web app
 $SiteCollections = Get-SPSite -WebApplication $WebAppURL -Limit All
 
 #Loop through all site collections
    foreach($Site in $SiteCollections)
    {
     Write-host "Scanning Site Collection:" $site.Url
  ###Check Whether the User is a Site Collection Administrator
     foreach($SiteCollAdmin in $Site.RootWeb.SiteAdministrators)
        {
      if($SiteCollAdmin.LoginName.EndsWith($ClaimsUserID,1))
      {
       "Site Collection, $($Site.RootWeb.Title), $($Site.RootWeb.Url), Site Collection Administrators Group, Site Collection Administrator" | Out-File $ReportPath -Append
      }    
    }
   
  #Get all webs
  $WebsCollection = $Site.AllWebs
  #Loop throuh each Site (web)
  foreach($Web in $WebsCollection)
  {
       if($Web.HasUniqueRoleAssignments -eq $True)
             {
     Write-host "Scanning Site:" $Web.Url
     
     #Get Permissions of the user on Web
     $WebPermissions = Get-PermissionInfo $ClaimsUserID $Web
      
     #Export Web permission data to CSV file - Append
     $WebPermissions |  Export-csv $ReportPath  -notypeinformation -Append
    }
     
    #Check Lists with Unique Permissions
    Write-host "Scanning Lists on $($web.url)..."
    foreach($List in $web.Lists)
    {
              if($List.HasUniqueRoleAssignments -eq $True -and ($List.Hidden -eq $false))
                 {
      #Get Permissions of the user on list
                        $ListPermissions = Get-PermissionInfo $ClaimsUserID $List
       
      #Export Web permission data to CSV file - Append
      $ListPermissions |  Export-csv $ReportPath -notypeinformation -Append      
     }
     
     #Check Folders with Unique Permissions
     $UniqueFolders = $List.Folders | where { $_.HasUniqueRoleAssignments -eq $True }                   
                    #Get Folder permissions
                    foreach($folder in $UniqueFolders)
        {
                        $FolderPermissions = Get-PermissionInfo $ClaimsUserID $folder
     
      #Export Folder permission data to CSV file - Append
      $FolderPermissions |  Export-csv $ReportPath -notypeinformation -Append   
                    }
     
     #Check List Items with Unique Permissions
     $UniqueItems = $List.Items | where { $_.HasUniqueRoleAssignments -eq $True }
                    #Get Item level permissions
                    foreach($item in $UniqueItems)
        {
                        $ItemPermissions = Get-PermissionInfo $ClaimsUserID $Item
       
      #Export List Items permission data to CSV file - Append
      $ItemPermissions |  Export-csv $ReportPath -notypeinformation -Append   
                    }
    }
  }
 }
 Write-Host Permission Report Generated successfully!
}
 
#Input Variables
$WebAppURL = "http://intranet.crescent.com"
$Userid ="Crescent\Salaudeen"
$ReportPath = "D:\Reports\PermissionRpt.csv"
 
#Call the function to generate user access report
Generate-PermissionReport $Userid $WebAppURL $ReportPath

