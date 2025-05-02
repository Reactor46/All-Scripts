$Subject = “User account created” # Message Subject 
$Server = “smtp.telus.net” # SMTP Server 
$From = “New.User.Created@aldergrovecu.ca” # From whom we are sending an e-mail 
$To = “jbattista@aldergrovecu.ca” # To whom we are sending

#$Pwd = ConvertTo-SecureString “password” -AsPlainText -Force #Sender account password (Warning! Use a very restricted account for the sender, because the password stored in the script will be not encrypted)
#$Cred = New-Object System.Management.Automation.PSCredential(“accountname” , $Pwd) #Sender account credentials 
$encoding = [System.Text.Encoding]::UTF8 #Setting encoding to UTF8 for message correct display

#Powershell command for filtering the security log about created user account event  

$Body=Get-WinEvent -FilterHashtable @{LogName=”Security”;ID=4720} | Select TimeCreated,@{n=”Account creator”;e={([xml]$_.ToXml()).Event.EventData.Data | ? {$_.Name -eq “SubjectUserName”} |%{$_.’#text’}}},@{n=”User Account”;e={([xml]$_.ToXml()).Event.EventData.Data | ? {$_.Name -eq “SamAccountName”}| %{$_.’#text’}}} | select-object -first 1 

#Sending an e-mail. 
Send-MailMessage -From $From -To $To -SmtpServer $Server -Body “$Body” -Subject $Subject -Credential $Cred -Encoding $encoding
