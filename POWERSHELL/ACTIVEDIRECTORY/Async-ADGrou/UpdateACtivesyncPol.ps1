######################################################################
#               Author: Vikas Sukhija(http://msexchange.me)
#               Date:- 03/16/2016
#		Reviewer:-
#               Description:- Change activesync policy based on 
#               a particular AD group.
######################################################################

$date1 = get-date -format d
$date1 = $date1.ToString().Replace("/","-")
$dir= ".\logs"
$limit = (Get-Date).AddDays(-30)

$logs = ".\Logs" + "\" + "Processed_" + $date1 + "_.log"

$smtpServer = "smtpserver"
$fromadd = "DoNotReply@labtest.com"
$email1 = "VikasS@labtest.com"

Start-Transcript -Path $logs

######Add Quest Shell & define attrib/ group value############

If ((Get-PSSnapin | where {$_.Name -match "Quest.ActiveRoles.ADManagement"}) -eq $null)
{
	Add-PSSnapin Quest.ActiveRoles.ADManagement
}

#######Add exchange Shell ##############################

If ((Get-PSSnapin | where {$_.Name -match "Microsoft.Exchange.Management.PowerShell.E2010"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}

$Asyncpol = "ActiveSync Policy"
$defAsyncpol = "Default ActiveSync Policy"

$group = "AD Group"

#################################################################

$groupmem = Get-QADGroupMember $group -sizelimit 0

$groupmem

$Statefile = "$($group)-Name.csv"


# If the file doesn't exist, create it

   If (!(Test-Path $Statefile)){  
                $groupmem | select Name | Export-csv $Statefile -NoTypeInformation 
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

        Write-host "$Asyncpol will be updated to "$_.Name"" -foregroundcolor green
       	$checkasync = Get-CASMailbox -Identity $_.Name
        if($checkasync.ActiveSyncEnabled -eq $true){
	Set-CASMailbox -Identity $_.Name -ActiveSyncMailboxPolicy $Asyncpol}}
	
	
        if($_.state -eq "Removed") {
        $userid = "$_.Name"
        Write-host "$Asyncpol will be removed from "$_.Name"" -foregroundcolor Red
	$checkasync = Get-CASMailbox -Identity $_.Name
        if($checkasync.ActiveSyncEnabled -eq $true){
	Set-CASMailbox -Identity $_.Name -ActiveSyncMailboxPolicy $defAsyncpol}}
	
      }

$groupmem | select Name | Export-csv $StateFile -NoTypeInformation

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
$msg.Subject = "Activesync Policy Script error"
$msg.Body = $error
$smtp.Send($msg)
$error.clear()
       }
  else

      {
    Write-host "no errors till now"
      }

$path = ".\logs\"
$limit = (Get-Date).AddDays(-30) #for log recycling

########################Recycle logs ######################################

Get-ChildItem -Path $path  | Where-Object {  
$_.CreationTime -lt $limit } | Remove-Item -recurse -Force 
stop-transcript

##########################################################################