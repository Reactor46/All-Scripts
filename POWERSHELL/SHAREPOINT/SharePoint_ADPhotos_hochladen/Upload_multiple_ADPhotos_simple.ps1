##############################################################################################
# 
# PowerShell Skript to simply upload multiple Photos to Active Directory with simple Logging
# and Prechecks before uploading
# 
# Author: Victor Bitschnau
# 
# Version 1.0
# 
##############################################################################################

#-------------------------------------PLEASE CHECK--------------------------------------------
#---------------------------------------------------------------------------------------------

$overwrite = "true" #If you want to overwrite existing pictures, set this variable to "true"

$pathToPictures = "H:\Install\Pics" #only path to folder

$maxPictureSize = "20000" # in byte

$pathToLog = "H:\Install\PicUpload.log" #with name of the file

$mySite = "http://mysite.sp2010.de" #MySite URL

#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------

#DateStamp
$Date=$(Get-Date -format g) | add-content $pathToLog

#Using hashtable for logs.
$errorlist=@{} #Log file list

$queryPics = Get-ChildItem $pathToPictures -ErrorAction Stop
 
foreach ($pic in $queryPics)
{
$username = $pic.basename
$dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$root = $dom.GetDirectoryEntry()
$search = [System.DirectoryServices.DirectorySearcher]$root
$search.Filter = "(&(objectclass=user)(objectcategory=person)(samAccountName=$username))"
$result = $search.FindOne()

#Check if user exists
if ($result -ne $null)
	{
	$user = $result.GetDirectoryEntry()
	$thumbnailPhoto = $user.thumbnailPhoto
	#Check if there is already a Picture
	if ($thumbnailPhoto.Value -eq $null)
		{
		#Check if picture size is less than maximum
		if($pic.length -le $maxPictureSize)
			{
			#Add picture to user property
			[byte[]]$jpg = Get-Content $pic.FullName -encoding byte
		 	$user.put("thumbnailPhoto",  $jpg )
		 	$user.setinfo()
			Write-Host $user.displayname "Uploaded picture successfully!" -ForegroundColor Green
			$errorlist.Add($username, "Sucess")
			}
		else
			{
			Write-Host $pic.Name "is too big in size! Max picture size is 10KB for each user." -ForegroundColor Red
			$errorlist.Add($username, "Size too big")
			}
		}
	else
		{
		if ($overwrite -eq $true)
			{
			#Add picture to user property
			[byte[]]$jpg = Get-Content $pic.FullName -encoding byte
			$user.put("thumbnailPhoto",  $jpg )
			$user.setinfo()
			Write-Host $user.displayname "Picture replaced!" -ForegroundColor Green
			$errorlist.Add($username, "Replaced")
			}
		else
			{
			Write-Host $user.displayname "has already stored a Picture in AD." -ForegroundColor Yellow
			$errorlist.Add($username, "Has already a picture")
			}
		}
	}
else 
	{
	Write-Host $username " Does not exist!" -ForegroundColor Red
	$errorlist.Add($username, "User not found")
	}
}

#Create thumbnails and update SharePoint Profile Photo Store
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Update-SPProfilePhotoStore -CreateThumbnailsForImportedPhotos $true -MySiteHostLocation $mySite

#Notification
Write-Host "#Note: Log files can be found in $pathtopictures called 'PicUpload.log'. `n"
 
#Add Error list to log file
 $errorlist | out-string | add-content $pathToLog
