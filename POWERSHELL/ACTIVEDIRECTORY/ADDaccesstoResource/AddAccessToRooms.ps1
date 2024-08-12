#############################################################################
#       Author: Vikas Sukhija
#       Reviewer:    
#       Date: 04/06/2015
#       Update: 
#	Description: Add group as Author, reviewer or other access to 
#	ALL Rooms/Equipment resources
# None 
# Owner
# PublishingEditor
# Editor
# PublishingAuthor 
# Author  
# NonEditingAuthor   
# Reviewer
# Contributor
#############################################################################
#########################Define Variables####################################
$access = "Author"     # Vaules can be anything from above
$usrgp =  "AllRoomsAuthorAccess"

$days = (get-date).adddays(-1)
$limit = (Get-Date).AddDays(-60) #for log recycling

$date = get-date -format d
$date = $date.ToString().Replace(“/”, “-”)
$time = get-date -format t

$time = $time.ToString().Replace(":", "-")
$time = $time.ToString().Replace(" ", "")

$path = ".\logs\"
$output1 = ".\logs\" + "Addaccess_" + $date + "_" + $time + "_.log"
$output2 = ".\logs\" + "Powershell_" + $date + "_" + $time + "_.log"


###################### Add Exchange Shell####################################

<#If ((Get-PSSnapin | where {$_.Name -match "Microsoft.Exchange.Management.PowerShell.E2010"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}#>

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ExchangeServer/PowerShell/ -Authentication Kerberos
import-pssession $session


start-transcript -Path $output2
########################Get all Rooms/Equipment mailboxes####################

$Resources = Get-Mailbox -resultsize unlimited | where{($_.RecipientTypeDetails -like "RoomMailbox") -or ($_.RecipientTypeDetails -like "EquipmentMailbox")}

if ($Resources){

$Resources | foreach{
if($_.WhenMailboxCreated -ge $days){
$Name = $_.primarysmtpaddress
$now = get-date -Format t
Write-host "Processing ...........$Name"  -foregroundcolor Green

Add-MailboxFolderPermission -identity ([string]$Name + ":\Calendar") -User $usrgp  -AccessRights $access

ADD-content $output1 "$now ...processed ..$Name"
 }
}
}

########################Recycle logs ######################################

Get-ChildItem -Path $path  | Where-Object {  
$_.CreationTime -lt $limit } | Remove-Item -recurse -Force 


stop-transcript
###########################################################################
