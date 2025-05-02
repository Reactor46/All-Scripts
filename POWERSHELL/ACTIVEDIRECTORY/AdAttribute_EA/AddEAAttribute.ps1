######################################################################
#               Author: Vikas Sukhija
#               Date:- 12/27/2015
#		Reviewer:-
#               Description:- Add EA Attribute to 
#               a particular AD group members.
######################################################################

$date1 = get-date -format d
$date1 = $date1.ToString().Replace("/","-")
$dir= ".\logs"
$limit = (Get-Date).AddDays(-30)

$logs = ".\Logs" + "\" + "Processed_" + $date1 + "_.log"

$smtpServer = "smtp.labtest.com"
$fromadd = "DoNotReply@labtest.com"
$email1 = "vikas@labtest.com"

Start-Transcript -Path $logs

######Add Quest Shell & define attrib/ group value############

If ((Get-PSSnapin | where {$_.Name -match "Quest.ActiveRoles.ADManagement"}) -eq $null)
{
	Add-PSSnapin Quest.ActiveRoles.ADManagement
}


$Attrbv = "EnableSync"  #Attribute Value

$group = "TestGroup1" #group Name

$Adattrbute = "extensionattribute1" #Ad attribute that will be updated

#################################################################

$groupmem = Get-QADGroupMember $group -sizelimit 0 -includedproperties $Adattrbute

$Statefile = "$($group)-Name.csv"


# If the file doesn't exist, create it

   If (!(Test-Path $Statefile)){  
                $groupmem | select Name,$Adattrbute | Export-csv $Statefile -NoTypeInformation 
                }

# Check Changes
$Changes =  Compare-Object $groupmem $(Import-Csv $StateFile) -Property Name | 
                Select-Object Name,
                    @{n='State';e={
                        If ($_.SideIndicator -eq "=>"){
                            "Removed" } Else { "Added" }
                        }
                    }

$Changes | foreach-object{
         
	if($_.state -eq "Added") {

        Write-host "$Attrbv will be added to "$_.Name"" -foregroundcolor green
	Set-QADUser -identity $_.Name -ObjectAttributes @{$Adattrbute  = $Attrbv}
	}
	
        if($_.state -eq "Removed") {
        $userid = "$_.Name"
        Write-host "$Attrbv will be removed from "$_.Name"" -foregroundcolor Red
	Set-QADUser -identity $_.Name -ObjectAttributes @{$Adattrbute  = $null}
	}
      }

$groupmem | select Name,$Adattrbute | Export-csv $StateFile -NoTypeInformation

###########################Recycle##########################################

$path = $dir 
 
Get-ChildItem -Path $path  | Where-Object {  
$_.CreationTime -lt $limit } | Remove-Item -recurse -Force 

#######################Report Error#########################################
if ($error -ne $null)
      {
#SMTP Relay address
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

#Mail sender
$msg.From = $fromadd
#mail recipient
$msg.To.Add($email1)
$msg.Subject = "AD Attribute Script Error"
$msg.Body = $error
$smtp.Send($msg)
$error.clear()
       }
  else

      {
    Write-host "no errors till now"
      }

$path = ".\logs\"
$limit = (Get-Date).AddDays(-60) #for log recycling

########################Recycle logs ######################################

Get-ChildItem -Path $path  | Where-Object {  
$_.CreationTime -lt $limit } | Remove-Item -recurse -Force 
stop-transcript

##########################################################################


