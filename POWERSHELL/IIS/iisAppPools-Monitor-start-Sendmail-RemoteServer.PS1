# DESCRIPTION: Script to monitor and start Application pools in multiple remote servers and get alert
# AUTHOR: Ripon Kundu - RIPONKUNDU@OUTLOOK.COM
# VERSION 1.0
# =======Application pool state========
# 0 > "Starting State"
# 1 > "Started State"
# 2 > "Stopping State"
# 3 > "Stopped State"
# 4 > "Unknown State"
# =====================================
#Invoke-Command can also do this, however this need Remote PowerShell execution to be enabled
# and that may be an risky permission to allow. Example of Invoke command to start app pool
#"Invoke-Command -ComputerName "Servername" -ScriptBlock { Start-WebAppPool -Name "Default Application" }"
#
#
#Change it to From address
 $from = "ApplicationMonitoring@contoso.com"
#Change it to Receipients address
 $to = "ripon.kundu@contoso.com"
#Replace this with SMTP server of your environment
 $smtphost = "smtprelay.corp.contoso.com"
#Add multiple server host names or IPs to be monitored
 $Servers = "SERVER01","SERVER02"
 Foreach ($Server in $Servers)
 {   
  #Get Application pool froom remote server
   $AppPools = gwmi -namespace "root\webadministration" -Class applicationpool -ComputerName $Server -Authentication PacketPrivacy -Impersonation Impersonate
   Foreach ($AppPool in $AppPools)
   {
    #Check Status of App Pool, Status returns 0 t0 4 and these refers to above descripted state
    if ($AppPool.GetState().ReturnValue -eq 4){
     sendEmail $AppPool.Name $Server "Unknown State"
    }
    elseif ($AppPool.GetState().ReturnValue -eq 3){
     sendEmail $AppPool.Name $Server "Stopped State and we starting it"
     $AppPool.start()
    }
    elseif ($AppPool.GetState().ReturnValue -eq 2){
     sendEmail $AppPool.Name $Server "Stopping State"
    }
    else{}
   }
 }

#Method for sendmail()
 Function sendEmail
 { param($AppPool,$Srv,$State)
#Email subject
 $subject = "$srv : $AppPool- $State"
#Email body
 $body = "$AppPool is in $State on $Srv, Please login to server for more info."
 $smtp= New-Object System.Net.Mail.SmtpClient $smtphost
 $msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
 $msg.isBodyhtml = $true
 $smtp.send($msg)
 }
