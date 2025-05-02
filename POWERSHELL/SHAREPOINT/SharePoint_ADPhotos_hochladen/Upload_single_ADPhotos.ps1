##############################################################################################
# 
# PowerShell Skript to manually upload single Photos to Active Directory
# 
# Author: Victor Bitschnau
# 
# Version 1.0
# 
##############################################################################################

#-------------------------------------PLEASE CHECK--------------------------------------------
#---------------------------------------------------------------------------------------------

$username = "<username>"
$jpgfile = "<path>"

#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------

$dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$root = $dom.GetDirectoryEntry()
$search = [System.DirectoryServices.DirectorySearcher]$root
$search.Filter = "(&(objectclass=user)(objectcategory=person)(samAccountName=$username))"
$result = $search.FindOne()
 
if ($result -ne $null)
{
 $user = $result.GetDirectoryEntry()
 [byte[]]$jpg = Get-Content $jpgfile -encoding byte
 $user.put("thumbnailPhoto",  $jpg )
 $user.setinfo()
 Write-Host $user.displayname "updated"
}
else {Write-Host $user "Does not exist"}
