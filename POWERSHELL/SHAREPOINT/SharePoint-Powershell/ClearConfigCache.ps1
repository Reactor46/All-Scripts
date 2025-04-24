Add-PSSnapin Microsoft.SharePoint.PowerShell –erroraction SilentlyContinue

#Stop SharePoint 2013 Timer service
Stop-Service SPTimerV4

#Configuration Cache folder
$folders = Get-ChildItem C:\ProgramData\Microsoft\SharePoint\Config 

#Get the GUID folder where cache.ini lives.
foreach ($folder in $folders)
    {
		$items = Get-ChildItem $folder.FullName -Recurse
		foreach ($item in $items)
        {
            if ($item.Name.ToLower() -eq "cache.ini")
                {
                    $cachefolder = $folder.FullName
                }                
        }
    }
	
#Delete all XML Files
$cachefolderitems = Get-ChildItem $cachefolder -Recurse
    foreach ($cachefolderitem in $cachefolderitems)
        {
            if($cachefolderitem -like "*.xml")
                {
                   $cachefolderitem.Delete()
                }        
        }
		
#Set Cache.ini files content to 1       
$a = Get-Content  $cachefolder\cache.ini
$a  = 1
Set-Content $a -Path $cachefolder\cache.ini

#Start SharePoint Timer Service
start-Service SPTimerV4

#Read more: https://www.sharepointdiary.com/2014/09/clear-sharepoint-2013-configuration-cache-using-powershell.html#ixzz6T2q3TDbA
