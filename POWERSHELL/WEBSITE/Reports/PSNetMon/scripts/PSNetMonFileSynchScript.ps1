﻿###################################################################################################
# PSNetMon - File Synch Script
# Written by Brad Voris
###################################################################################################
#To change the amount of time that the files are scrutinized change the number of addminutes in the $time variable
#Keep in mind the frequency that this script will run in the task scheduler could put timing slightly off
#Modify the path variable to the correct path
#Variables
$Time = (Get-Date).AddMinutes(-10)
$CSS = get-content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\css\theme.css"
$Dated = (Get-Date -format F)
$Path = "C:\folder"
#Expression
$Data = Get-ChildItem -Path $Path -Recurse -Force | ` Where-Object {$_.LastWriteTime -le $Time} | Select FULLNAME, LastWriteTime | ConvertTo-Html
#HTML Head
$HTMLHead = @"
<!DOCTYPE html>
<HEAD>
<META charset="UTF-8">
<TITLE>PSNetMon - File Synch  Module</TITLE>
<CENTER>
<STYLE>$CSS</STYLE>
"@
#HTML Body Content
$HTMLBody = @"
<Font size=4><B>File Synch Module</B></font></BR>

#Export to HTML