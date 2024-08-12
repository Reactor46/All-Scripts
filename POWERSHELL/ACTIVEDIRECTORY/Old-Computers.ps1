###############################################################
#<#foreach {Set-ADComputer -Identity $_ -Enabled $false $_}|#>#
###############################################################

$DaysInactive = 180 
$time = (Get-Date).Adddays(-($DaysInactive))
Get-ADComputer -filter {((LastLogonTimeStamp -lt $time) -or (LastLogonTimeStamp -notlike "*"))} -SearchBase "DC=contoso,DC=com" -Properties CN,name,LastLogonDate,*|        
sort LastLogonDate -Descending | ft name,ipv4address,LastLogonDate | FT -AutoSize