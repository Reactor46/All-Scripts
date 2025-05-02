# ** THIS SCRIPT IS PROVIDED WITHOUT WARRANTY, USE AT YOUR OWN RISK **	

# DESCRIPTION
# This script gets system event log errors for the last 5 days. I have it set to dump the ouput file in
# HTML format to a share drive location.

# You can create a scheduled task on a server by calling this script via a bat file.

# REQUIREMENTS
# 1.	Set-Executionpolicy remotesigned

# NOTES
# Tested with Windows 7, Windows Vista, Windows Server 2003, Windows Server 2K8 and 2K8 R2

# AUTHOR
# David Hall | http://www.signalwarrant.com/

# LINK

$logPath = "C:\inetpub\wwwroot\Report\ServerEventLogs\"
$log = "system"
$computers = "LASDC01" #GC -Path C:\inetpub\AdminScripts\Configs\All-Servers.txt

# Start HTML Output file style
$style = "<style>"
$style = $style + "Body{background-color:white;font-family:Arial;font-size:10pt;}"
$style = $style + "Table{border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}"
$style = $style + "TH{border-width: 1px; padding: 2px; border-style: solid; border-color: black; background-color: #cccccc;}"
$style = $style + "TD{border-width: 1px; padding: 5px; border-style: solid; border-color: black; background-color: white;}"
$style = $style + "</style>"

# End HTML Output file style

$date = get-date -format M.d.yyyy

$now = get-date
$subtractDays = New-Object System.TimeSpan 5,0,0,0,0
$then = $Now.Subtract($subtractDays)

$systemErrors = Get-EventLog -Computername $computers -LogName $log -After $then -Before $now -EntryType Error |
select EventID,MachineName,Message,Source,TimeGenerated

$systemErrors | ConvertTo-HTML -head $style | Out-File "$logPath\$computers-$log-$date.htm"