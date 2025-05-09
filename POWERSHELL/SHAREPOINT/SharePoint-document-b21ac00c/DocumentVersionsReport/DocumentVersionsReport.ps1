[bool] $ReportOnlyFilesWithVersions = $true

#write header
$OutFile = "C:\Scripts\StorageReport-" + $(get-date -f yyyyMMdd-HHmmss) + ".txt"
$file = New-Object System.IO.StreamWriter $OutFile

$file.WriteLine("WebApplication`tSiteCollection`tSite`tDocumentLibrary`tItemName`tItemVersions`tItemSize`tItemVersionsSize`tItemTotalSize") 

$SiteUrl = "http://teams.corp.local"
$site = new-Object Microsoft.SharePoint.SPSite($SiteUrl)           
$webapp = $site.WebApplication
foreach($currentsite in $webapp.Sites)
{
   write-host $currentSite.Url
   foreach($currentWeb in $currentSite.AllWebs)
   {              
      write-host $currentWeb.Url
      $lists = $currentWeb.Lists | Where {$_.BaseType -eq "DocumentLibrary"} 
      foreach($currentList in $lists)
      {
         foreach($currentListItem in $currentList.Items)                 
         {  
           $versionssize = 0
           if($currentListItem.Versions.Count -gt 1)
           {
                foreach($version in $currentListItem.File.Versions)
                {
                   $versionssize = $versionssize + $version.Size
                }
                $totalsize = $versionssize + $currentListItem.File.Length
                $line = "{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}`t{8}" -f $webapp.Url,$currentSite.ServerRelativeUrl,$currentWeb.ServerRelativeUrl,$currentList.Title, $currentListItem.Name, $currentListItem.Versions.Count, $currentListItem.File.Length, $versionssize ,$totalsize 
                $file.WriteLine($line)
           
           }
           else
           {
                if ($ReportOnlyFilesWithVersions -eq $false)
                { 
                    $totalsize = $versionssize + $currentListItem.File.Length
                    $line = "{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}`t{8}" -f $webapp.Url,$currentSite.ServerRelativeUrl,$currentWeb.ServerRelativeUrl,$currentList.Title, $currentListItem.Name, $currentListItem.Versions.Count, $currentListItem.File.Length, $versionssize ,$totalsize 
                    $file.WriteLine($line)
                }
           
           }
                      
         }
      }                                    
      $currentWeb.Dispose()         
   }                 
   $currentSite.Dispose()
}
$site.Dispose()  

$file.Close()
 