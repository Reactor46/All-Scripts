# Originally published at https://gist.github.com/nzthiago/5736907 
# and also Vlad have updload Powershell script on technet gallery to download SPC2014 videos & slides http://gallery.technet.microsoft.com/PowerShell-Script-to-all-04e92a63
# I Customized it to download video series of SharePoint Hybrid Courseware and Curriculum
# If you like it, leave me a comment

[Environment]::CurrentDirectory=(Get-Location -PSProvider FileSystem).ProviderPath 
$rss = (new-object net.webclient)


# SharePoint Hybrid courseware & curriculum series Videos
$a = ([xml]$rss.downloadstring("http://channel9.msdn.com/series/SharePoint-Hybrid-Courseware-and-Curriculum/RSS/mp4high")) 

#other qualities for the videos only. Choose the one you want!
# $a = ([xml]$rss.downloadstring("http://channel9.msdn.com/series/SharePoint-Hybrid-Courseware-and-Curriculum/RSS/mp4")) 
#$a = ([xml]$rss.downloadstring("http://channel9.msdn.com/series/SharePoint-Hybrid-Courseware-and-Curriculum/RSS/mp3")) 


#Preferably enter something not too long to not have filename problems! cut and paste them afterwards
$downloadlocation = "C:\SharePointHybridCourseware"

	if (-not (Test-Path $downloadlocation)) { 
		Write-Host "Folder $fpath dosen't exist. Creating it..."  
		New-Item $downloadlocation -type directory 
	}
set-location $downloadlocation

# Walk through each item in the feed 
$a.rss.channel.item | foreach{   
	$code = $_.comments.split("/") | select -last 1	   
	
	# Grab the URL for the MP4 file
	$url = New-Object System.Uri($_.enclosure.url)  
	
	# Create the local file name for the MP4 download
    $file = $_.creator + "-" + $_.title.Replace(":", "-").Replace("?", "").Replace("/", "-").Replace("<", "").Replace("|", "").Replace('"',"").Replace("*","")
	$file = $file.substring(0, [System.Math]::Min(120, $file.Length))
	$file = $file.trim()
	$file = $file + ".mp4"  

    #Creating Folder	
	if ($code -ne "")
	{
         $folder = $_.title.Replace(":", "-").Replace("?", "").Replace("/", "-").Replace("<", "").Replace("|", "").Replace('"',"").Replace("*","")
		 $folder = $folder.substring(0, [System.Math]::Min(100, $folder.Length))
		 $folder = $folder.trim()
	}
	else
	{
		$folder = "NoCodeSessions"
	}
	
	if (-not (Test-Path $folder)) { 
		Write-Host "Folder $folder) dosen't exist. Creating it..."  
		New-Item $folder -type directory 
	}
	
	# Make sure the MP4 file doesn't already exist

	if (!(test-path "$folder\$file"))     
	{ 	
		# Echo out the  file that's being downloaded
		$file
		$wc = (New-Object System.Net.WebClient)  

		# Download the MP4 file
		$wc.DownloadFile($url, "$downloadlocation\$file")
		mv $file $folder
	}
	
		
	}