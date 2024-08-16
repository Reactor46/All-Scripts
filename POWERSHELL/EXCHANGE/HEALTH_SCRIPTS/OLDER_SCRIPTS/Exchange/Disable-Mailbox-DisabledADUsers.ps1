
#Just check if is not ActiveDirectory module already imported. If not,import it 
if (!(get-module ActiveDirectory)) 
{ 
    Write-Output "Importing ActiveDirectory module" 
    Import-Module activedirectory 
} 
 
#Just check if is not Exchange 2013/16 snapin already imported. If not,import it 
if(!(Get-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) 
{ 
    write-Output "Importing Exchange 2013/16 snapin" 
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 
} 
 
#Parameters for logging purpose 
$PsDate=(get-date -Format dd.MM.yyyy-HH:mm) 
$PSLogPath=($MyInvocation.InvocationName -replace ".ps1", ".log") 
$ScriptName=($MyInvocation.MyCommand -replace ".ps1", "") 
 
#SMTP settings: 
$SMTPserver="server@domain.sk" 
$SentFrom=$env:ComputerName+"@domain.sk" 
$SentTo="administrator@domain.sk" 
$EmailSubject="Notification "+ $ScriptName 
$EmailHeader=$PsDate 
$EmailBody="" 
$EmailBodyFooter="`n`nFOOTER:`n"+$MyInvocation.InvocationName+" `n"+$PSLogPath 
 
#Getting Disabled users and user with attribute msExchRecipientTypeDetails=1 which means UserMailbox (not Room mailbox, not others...) 
$DisabledUsers=Get-ADUser -filter {Enabled -eq $False -and msExchRecipientTypeDetails -eq 1}  
$DisabledUsersCount=($DisabledUsers | measure-object).count 
 
#if disabled users count is greater than 0, continue: 
if($DisabledUsersCount -gt 0) 
{ 
    ForEach($DisabledUser in $DisabledUsers) 
    { 
        Disable-Mailbox -identity $DisabledUser.SamAccountName -confirm:$FALSE  -ErrorVariable +Errors 
        #give all user names which has been disabled - logging purpose: 
        $EmailBody+="`n"+$($DisabledUser.Name) 
    } 
 #Save all logging information into file:    
 Add-Content -path $PSLogPath -value $($PsDate + $EmailBody) 
 #send all logging information into administrator email 
 Send-MailMessage -smtpserver $SMTPserver -from $SentFrom -to $SentTo -Subject $EmailSubject -Body $($EmailHeader + " `n" + $ScriptName + " Count: " + $DisabledUsersCount +"`n"+ $EmailBody + $EmailBodyFooter +"`n"+ $errors) 
} 