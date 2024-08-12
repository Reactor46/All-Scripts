<#
*** THIS SCRIPT IS PROVIDED WITHOUT WARRANTY, USE AT YOUR OWN RISK ***

.DESCRIPTION
	1. Search an OU for user accounts that have not authenticated in x number of days ($days)
    2. Disable those accounts
    3. Move those disabled user accounts to another OU ($disabledOU)
    4. Also creates a logfile of all the userss that were disabled ($logpath)

.NOTES
	File Name: Get-StaleUsers.ps1
	Author: David Hall
	Contact Info: 
		Website:&nbsp;www.signalwarrant.com
		Twitter:&nbsp;@signalwarrant
		Facebook:&nbsp;facebook.com/signalwarrant/
		Google +:&nbsp;plus.google.com/113307879414407675617
		YouTube Subscribe link: https://www.youtube.com/channel/UCgWfCzNeAPmPq_1lRQ64JtQ?sub_confirmation=1
	Requires: PowerShell Remoting Enabled (Enable-PSRemoting) 
	Tested: PowerShell Version 5, Windows Server 2012 R2


    This script will search an Organizational Unit for Users accounts that have not authenticated to the Domain in 1 hour.
    You can easily modify the number of hours or change it to days by replacing this bit of code.

    # Change this on line 66
        Today.AddHours($hours)
 
    # to this on line 66
        Today.AddDays($hours)

.PARAMETER
		 
.EXAMPLE
     .\Get-StaleUsers.ps1
#>

###################################
####### Edit these Variables
###################################

# Gets todays Date
$date = Get-Date

# Number of days it's been since the computer authenticated to the domain
# In my case 1 day
$hours = "-1"

# Sets a description on that object so other admins know why the object was disabled
$description = "Disabled by SignalWarrant on $date due to inactivity for 1 days."

# This is the OU you are searching for Stale Computer accounts
$ou = "DC=contoso,DC=com"

# This is where the disabled accounts get moved to.
$disabledOU = "OU=UserAccounts,OU=Disabled,DC=contoso,DC=com"

# path to the log file
$logpath = "c:\scripts\disabled_users.csv"

###################################
#######
###################################

$finduser = Get-aduser –filter * -SearchBase $ou -properties cn,lastlogondate | 
Where {$_.LastLogonDate –le [DateTime]::Today.AddHours($hours) -and ($_.lastlogondate -ne $null) }

$finduser | export-csv $logpath
$finduser | set-aduser -Description $description –passthru | Disable-ADAccount

write-host -foregroundcolor Green "Searching OU for disabled User Accounts"
[System.Threading.Thread]::Sleep(500)

$disabledAccounts = Search-ADAccount -AccountDisabled -UsersOnly -SearchBase $ou

write-host -foregroundcolor Green "Moving disabled Users to the Disabled_Users OU"
[System.Threading.Thread]::Sleep(500)

$disabledAccounts | Move-ADObject -TargetPath $disabledOU

write-host -foregroundcolor Green "Script Complete"