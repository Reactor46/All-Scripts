param (
    [Parameter(Mandatory=$true)]
    [string]$importfile = "",
    [string]$logpath = ""
    )

#Check SharePoint major build version.    
[Void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$SPFarm = [Microsoft.SharePoint.Administration.SPfarm]::Local
$version = $SPFarm.BuildVersion.Major

#Set up the STSADM command path regarding what is the major build version
$stsadmpath = "c:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\$version\BIN\"

#Set location where we are.
$location = Get-Location
Set-Location $location

#Create a log folder if we signed, anyway set it default location.
if ($logpath -ne "") 
    {New-Item -ItemType directory -Path $logpath -erroraction SilentlyContinue | 
    Write-Host > $null }
    else
    {$logpath += $location}

#Start working...
$title = "Work with SP sites"
$message = "What do you want to do with the site collection?"

#Create all choice for jobs as backup, delete, restore and quite.
$backup = New-Object System.Management.Automation.Host.ChoiceDescription "&Backup", `
    "Backup all site collection from source list."

$delete = New-Object System.Management.Automation.Host.ChoiceDescription "&Delete", `
    "Delete all site collection from source list."
    
$restore = New-Object System.Management.Automation.Host.ChoiceDescription "&Restore", `
    "Restore all site collection from source list."

$quit = New-Object System.Management.Automation.Host.ChoiceDescription "&Quit", `
    "Quit."    
   
#Initiate the menu.
$options = [System.Management.Automation.Host.ChoiceDescription[]]($backup, $delete, $restore, $quit)

#Do the jobs until you want to quit.
:OuterLoop do
{
 #Initiate a prompt for choice.
 $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
 
 #Do the job what you choose.
 switch ($result)
  {	
   #Create a backup from all site collection what are in the .csv file.
   0 {"BACKUP"
     $DateStr = "{0:ddMMyyyyhhmmss}" -f (get-date)
     $stsadmpath
     Write-Host -ForegroundColor green "Opening SharePoint site..."

     ipcsv $importfile | 
     foreach {
        Set-Location $stsadmpath
        $site = $_.Site
        $file = "$location\" + $_.Site.Replace("http://","").Replace("https://","").Replace("/","_").Replace(";","_") + ".bak"
        $result = .\stsadm.exe -o backup -url $site -filename $file -overwrite
        $DateStrForLog = "{0:dd/MM/yyyy hh:mm:ss}" -f (get-date)
        if ($result -eq "Operation completed successfully.")
            {
            write-host -ForegroundColor magenta "[$DateStrForLog] " -nonewline
            write-host -ForegroundColor green  "The " -nonewline 
            write-host -ForegroundColor magenta $site -nonewline 
            write-host -ForegroundColor green " has been backed up to " -nonewline
            write-host -ForegroundColor yellow $file -nonewline
            write-host -ForegroundColor green `n $result
            $logentry = "[$DateStrForLog] The $site has been backed up to `"$file`"" | 
            out-file $logpath\$DateStr"-backup.log" -Append
            $result | out-file $logpath\$DateStr"-backup.log" -Append
            }
            else
            {
            write-host -ForegroundColor magenta "[$DateStrForLog] " -nonewline
            write-host -ForegroundColor green "The " -nonewline
            write-host -ForegroundColor red $site -nonewline
            write-host -ForegroundColor green " site could not be found!" `n $result
            $logentry = "[$DateStrForLog] The $site site could not be found!" | 
			out-file $logpath\$DateStr"-backup.log" -Append
            $result = "" |  out-file $logpath\$DateStr"-backup.log" -Append                        
            $result = "The site not found!" | 
            out-file $logpath\$DateStr"-backup.log" -Append
            $result = "" |  out-file $logpath\$DateStr"-backup.log" -Append
            $site | out-file $logpath\$DateStr"-sitesnotfound.log" -Append                        
	        }
        }
        Write-Host -ForegroundColor green "Done."
        Set-Location $location
     }
		
   #Delete all site collection what are in the .csv file
   1 {"DELETE"
     $DateStr = "{0:ddMMyyyyhhmmss}" -f (get-date)
     $stsadmpath
     Write-Host -ForegroundColor green "Opening SharePoint site..."

     ipcsv $importfile | 
     foreach {
        Set-Location $stsadmpath
        $site = $_.Site
        $result = .\stsadm.exe -o deletesite -url $site
        $DateStrForLog = "{0:dd/MM/yyyy hh:mm:ss}" -f (get-date)
        if ($result -eq "Operation completed successfully.")
        	{
            write-host -ForegroundColor magenta "[$DateStrForLog] " -nonewline
            write-host -ForegroundColor green  "The " -nonewline 
            write-host -ForegroundColor magenta $site -nonewline
            write-host -ForegroundColor green " has been deleted! " `n $result 
            $logentry = "[$DateStrForLog] The $site has been deleted!" | 
            out-file $logpath\$DateStr"-delete.log" -Append
            $result | out-file $logpath\$DateStr"-delete.log" -Append
            }
            else
            {
            write-host -ForegroundColor magenta "[$DateStrForLog] " -nonewline
            write-host -ForegroundColor green "The " -nonewline
            write-host -ForegroundColor red $site -nonewline
            write-host -ForegroundColor green " site could not be found!" `n $result
            $logentry = "[$DateStrForLog] The $site site could not be found! $result" | 
            out-file $logpath\$DateStr"-delete.log" -Append
            $result = "" |  out-file $logpath\$DateStr"-delete.log" -Append
            $result = "The site not found!" |  
			out-file $logpath\$DateStr"-delete.log" -Append
            $result = "" |  out-file $logpath\$DateStr"-delete.log" -Append
            $site | out-file $logpath\$DateStr"-sitesnotfound.log" -Append
            }
        }
        write-host -ForegroundColor green "Done."
        Set-Location $location
     }
		  
	#Restore all site collection what are in the .csv file. 
	#WARNING: Check the connection of sites and backups!  
    2 {"RESTORE"
      $DateStr = "{0:ddMMyyyyhhmmss}" -f (get-date)
      $stsadmpath
      Write-Host -ForegroundColor green "Opening SharePoint site..."

      ipcsv $importfile | 
      foreach {
        Set-Location $stsadmpath
        $site = $_.Site
        $file = "$location\" + $_.Site.Replace("http://","").Replace("https://","").Replace("/","_").Replace(";","_") + ".bak"
        $DateStrForLog = "{0:dd/MM/yyyy hh:mm:ss}" -f (get-date)
        if ((Test-Path $file) -eq $True)
        	{
            $result = .\stsadm.exe -o restore -url $site -filename $file -overwrite
            $DateStrForLog = "{0:dd/MM/yyyy hh:mm:ss}" -f (get-date)
            if ($result -eq "Operation completed successfully.")
            	{
                .\stsadm.exe -o refreshsitedms -url $site | out-null
                .\stsadm.exe -o setsitelock -url $site -lock none | out-null
                write-host -ForegroundColor magenta "[$DateStrForLog] " -nonewline
                write-host -ForegroundColor green  "The " -nonewline 
                write-host -ForegroundColor magenta $site -nonewline 
                write-host -ForegroundColor green " has been restored from " -nonewline 
                write-host -ForegroundColor yellow $file -nonewline
                write-host -ForegroundColor green `n $result
                $logentry = "[$DateStrForLog] The $site has been restored from `"$file`"" | 
                out-file $logpath\$DateStr"-restore.log" -Append
                $result | out-file $logpath\$DateStr"-restore.log" -Append
                }
                else
                {
                write-host -ForegroundColor magenta "[$DateStrForLog] " -nonewline
                write-host -ForegroundColor green "The " -nonewline
                write-host -ForegroundColor red $site -nonewline
                write-host -ForegroundColor green " site could not be found!" `n $result
                $logentry = "[$DateStrForLog] The $site site has not been restored correctly! $result" | 
                out-file $logpath\$DateStr"-restore.log" -Append
                $result = "" |  out-file $logpath\$DateStr"-restore.log" -Append
                $result = "The site has not been restored correctly!" |  
				out-file $logpath\$DateStr"-restore.log" -Append
                $result = "" |  out-file $logpath\$DateStr"-restore.log" -Append
                $site | out-file $logpath\$DateStr"-sitesnotfound.log" -Append
                }
            }
            else
            {
            write-host -ForegroundColor magenta "[$DateStrForLog] " -nonewline
            write-host -ForegroundColor green "The " -nonewline
            write-host -ForegroundColor red $file -nonewline
            write-host -ForegroundColor green " backup file could not be found!"`n
            $result = "" |  out-file $logpath\$DateStr"-restore.log" -Append
            $logentry = "[$DateStrForLog] The $file backup file could not be found!" |  
            out-file $logpath\$DateStr"-restore.log" -Append
            $result = "" |  out-file $logpath\$DateStr"-restore.log" -Append
            }   
        }
        Write-Host -ForegroundColor green "Done."
        Set-Location $location
      }
		
	#Quit from job.
    3 {"QUIT"
      Set-Location $location;
      break OuterLoop
      }
  }
}
while ($y -ne 100)