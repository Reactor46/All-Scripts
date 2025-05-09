##############################################################################################
# 
# PowerShell Skript to upload multiple Photos to Active Directory with automatic Logging
# management and Prechecks before uploading
# 
# Author: Victor Bitschnau
# 
# Version 1.0
# 
##############################################################################################

$Date = $(Get-Date -Uformat "%d.%m.%Y-%H.%M")

#-------------------------------------PLEASE CHECK--------------------------------------------
#---------------------------------------------------------------------------------------------

$overwrite = "true" #If you want to overwrite existing pictures, set this variable to "true"

$pathToPictures = "H:\Install\Pics" #only path to folder

$uploadedFolder = "H:\Install\Uploaded" #Folder for uploaded pictures

$errorFolder = "H:\Install\ErrorPics"

$maxPictureSize = "20000" # in byte

$pathToLog = "H:\Install\PicUpload.log" #with name of the file

$maxLogSize = "10000000" #in byte

$mySite = "http://mysite.sp2010.de" #MySite URL

$backupFolder = "H:\Install\Logs\"  #Backup-File Location

$backupDays = "7" #Days until a Backup file will be deleted

#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------

$log = Get-Item $pathToLog

#Check Log file size
if ($log.length -gt $maxLogSize)
{
$backupLog = "$backupFolder\PicUpload-$Date.log"
Copy-Item $pathToLog $backupLog
Clear-Content $pathToLog
Write-Host "Log file copied to $backupLog" -ForegroundColor Green
}

#Delete old Backup files in Backup folder

$backupFiles = Get-ChildItem h:\Install\Logs

#Find the old backups
$old = $backupFiles | ? { $_.CreationTime -lt ((get-date).adddays(-$backupDays)) }
if ($old -eq $Null) 
	{ 
	Write-Host "No backups older than $backupDays days found in '$backupFolder'!"
	}
#Delete the old Backup files
else 
	{
	$old | % { Remove-Item $backupFolder$_ }
	Write-Host "Deleted old Backup files!"
	}

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
			Move-Item $pic.FullName -Destination $uploadedFolder
			Write-Host $user.displayname "Uploaded picture successfully and moved to '$uploadedFolder' folder!" -ForegroundColor Green
			$errorlist.Add($username, "Sucess")
			}
		else
			{
			Write-Host $pic.Name "is too big in size! Max picture size is 10KB for each user. Picture was moved to '$errorFolder' folder!" -ForegroundColor Red
			Move-Item $pic.FullName -Destination $errorFolder
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
			Move-Item $pic.FullName -Destination $uploadedFolder
			Write-Host $user.displayname "Picture replaced and moved to '$uploadedFolder' folder!" -ForegroundColor Green
			$errorlist.Add($username, "Replaced")
			}
		else
			{
			Write-Host $user.displayname "has already stored a Picture in AD. Picture was moved to '$errorFolder' folder!" -ForegroundColor Yellow
			Move-Item $pic.FullName -Destination $errorFolder
			$errorlist.Add($username, "Has already a picture")
			}
		}
	}
else 
	{
	Write-Host $username " Does not exist! Picture was moved to '$errorFolder' folder!" -ForegroundColor Red
	Move-Item $pic.FullName -Destination $errorFolder
	$errorlist.Add($username, "User not found")
	}
}

#Create thumbnails and update SharePoint Profile Photo Store
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Update-SPProfilePhotoStore -CreateThumbnailsForImportedPhotos $true -MySiteHostLocation $mySite

#Notification
$maxLogSize = $maxLogSize/1000000
Write-Host "#Note: Log files can be found in $pathtopictures called 'PicUpload.log' and will be copied to 'H:\Install\Logs\PicUpload-Date-Time.log' after size exceeds $maxLogSize MB. `n"
 
#Add Error list to log file
 $errorlist | out-string | add-content $pathToLog
 
