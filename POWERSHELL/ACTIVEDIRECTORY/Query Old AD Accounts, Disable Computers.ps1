Import-Module ActiveDirectory  

$DaysInactive = 45
$time = (Get-Date).Adddays(-($DaysInactive)) 

# This script is to notify us of inactive user accounts.

$html_open = 
"
<!DOCTYPE html>`n
<html>`n
<head>`n
    <style>`n
    body {`n
    font: 12px Arial, sans-serif;`n
    }`n
    h4 {`n
        color: #000080;`n
    }`n
    </style>`n
</head>`n
<body>`n
"

$html_close =
"
</body>`n
</html>`n
"

$table_open =
'
<table style ="' + 'width:100% + ">'

$table_close = '</table><br><br>'

$decs = ""
$details = ""
$output = ""
$results = ""
$message = ""

$smtp = "smtp.office365.com"
$user = "jack@netlinkinc.net"
$pass = "Friendship20..." | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $pass
$email = "jack@netlinkinc.net"

#Header Information
$today = (Get-Date).Date
$desc = "<h4>Inactive Active Direcotry Computer Accounts:</h4><hr>"

  
# Get all old AD computers
#Write-Host "Inactive Active Directory Computer Accounts:`n`n"
Get-ADComputer -Filter {LastLogonDate -lt $time -and Enabled -eq "true"} -Properties * | ForEach-Object {

    $output = "<b>" + ($_.Name | Out-String) + "</b>     has been inactive since     <b>" + ($_.LastLogonDate | Out-String) + "</b><br>"
    $results = $results + $output
}
$details = $details + "<br>" + $results
$details = $details + "<br><br><b><u>These accounts will be automatically disabled.</b></u>"
#Write-Host "`n`nThese accounts will be automatically disabled."

$message = $desc + $table_open + $details + $table_close

$details = ""

#Get all old AD users
$desc = "<h4>Inactive Active Directory User Accounts:</h4><hr>"
#Write-Host "Inactive Active Directory User Accounts:`n`n"

$results = ""
Get-ADUser -Filter {LastLogonDate -lt $time} -Properties * | ForEach-Object {

    $output = "<b>" + ($_.Name | Out-String) + "</b>     has been inactive since     <b>" + ($_.LastLogonDate | Out-String) + "</b><br>"
    $results = $results + $output
}
$details = $details + "<br>" + $results
$details = $details + "<br><br>These accounts will stay active until manual action is taken."

$message = $message + $desc + $table_open + $details + $table_close

# Disables all old AD computers
Get-ADComputer -Filter {LastLogonDate -lt $time} -Property LastLogonDate | Set-ADComputer -Enabled $false -Confirm:$False
#Get-ADUser -Filter {LastLogonDate -lt $time} -Property LastLogonDate | Set-ADUser -Enabled $false

$message = $html_open + $message + $html_close
Send-MailMessage -To "jack@netlinkinc.net" -Subject "Inactive AD Accounts" -Body $message.Trim() -BodyAsHtml -From $email -SmtpServer $smtp -usessl -Credential $cred