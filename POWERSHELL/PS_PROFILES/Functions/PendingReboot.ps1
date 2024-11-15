
#Notes
<#
This script checks through a list of computers to report via email
whether any computers are in a "Reboot Pending" state.
#>

#Email Variables
$smtpServer = "smtp.contoso.com"
$smtpFrom = "Reboot Pending Report <Reboot.Pending@contoso.com>"
$smtpTo = "Server Admin <server.admin@contoso.com>"
$Subject = "'Reboot Pending' report"

#Server List
<#Three typical ways to get a list of computers.
1. Comma separated manually entered list.
$CommaList = ("Server01","Server02")
$List = $CommaList.Split(",")
2. List OU members 
$List = Get-ADComputer -Filter * -SearchBase "OU=IT,DC=contoso,DC=com"
3. List from TXT file
$List = Get-Content -Path "c:\scripts\PC-List.txt"
#>

# Choose a List method as shown above and replace the following two lines.
#$CommaList = ("Server01","Server02")
$CommaList = ("usonvsvritsw.USON.LOCAL")
$List = $CommaList.Split(",") 


<#~~~~~~~~~~~~~~~ DO NOT EDIT BEYOND THIS POINT ~~~~~~~~~~~~~~~#>

#Table Style
$style = "<style>BODY{font-family: Calibri; font-size: 10pt;}"
$style = $style + "TABLE{border: 0px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 0px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 0px solid black; padding: 5px; }"
$style = $style + "</style>"
$THstyle = "font-size:16pt;font-weight:bold;"
$TDstyle = "font-weight:bold;text-align:center;"

#Clean Start 
$Checks = ("WUboot,WUVal,PakBoot,PakVal,PakWowBoot,PakWowVal,RenFileBoot,PCnameBoot,CBSboot,CBSVal,Content")
$Clear = $checks.split(",")
Clear-Variable -Name $clear

#Table header row
$Content =""
$Content += "<table id=""t1"">"
$Content += "<tr bgcolor=#ADD8E6>"
$Content += "<td width=100 style=$THstyle> Server</td>"
$Content += "<td width=75 style=$TDstyle>Reboot Required?</td>"
$Content += "<td width=75 style=$TDstyle>Windows Updates</td>"
$Content += "<td width=75 style=$TDstyle>Package Installer</td>"
$Content += "<td width=75 style=$TDstyle>Package Installer 64</td>"
$Content += "<td width=75 style=$TDstyle>File Rename</td>"
$Content += "<td width=75 style=$TDstyle>Hostname Change</td>"
$Content += "<td width=80 style=$TDstyle>Component Based Svces</td>"


foreach ($PC in $CommaList)
#List from OU  - Change searchbase from Split to OU
#List from TXT
{
$Content += "<tr>"

#Windows Updates
$WUVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\" -Name RebootRequired -ErrorAction SilentlyContinue}) | foreach { $_.RebootRequired }
if ($WUVal -ne $null) {$WUBoot = "Yes"}
else {$WUBoot = "No"}

#Package Installer
$PakVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\Software\Microsoft\Updates\" -Name UpdateExeVolatile -ErrorAction SilentlyContinue}) | foreach { $_.UpdateExeVolatile }
if ($PakVal -ne $null) {$PakBoot = "Yes"}
else {$PakBoot = "No"}


#Package Installer - Wow64
$PakWowVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\Software\Wow6432Node\Microsoft\Updates\" -Name UpdateExeVolatile -ErrorAction SilentlyContinue}) | foreach { $_.UpdateExeVolatile }
if ($PakWowVal -ne $null) {$PakWowBoot = "Yes"}
else {$PakWowBoot = "No"}


#Pending File Rename Operation
$RenFileVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue}) | foreach { $_.PendingFileRenameOperations }
if ($RenFileVal -ne $null) {$RenFileBoot = "Yes"}
else {$RenFileBoot = "No"}


#Pending Computer Rename
$PCnameIs = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" -Name ComputerName -ErrorAction SilentlyContinue}) | foreach { $_.ComputerName }
$PCnameBe = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -Name ComputerName -ErrorAction SilentlyContinue}) | foreach { $_.ComputerName }
if ($PCnameIs -eq $PCnameBe) {$PCnameBoot = "No"}
else {$PCnameBoot = "Yes"}


#Component Based Servicing
$CBSVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\" -Name RebootPending -ErrorAction SilentlyContinue}) | foreach { $_.RebootPending }
if ($CBSVal -ne $null) {$CBSBoot = "Yes"}
else {$CBSBoot = "No"}


#Email HTML Content - append loop
$Content += "<td bgcolor=#dddddd align=left><b>$PC</b></td>"
if (($WUboot,$PakBoot,$PakWowBoot,$RenFileBoot,$PCnameBoot,$CBSBoot) -contains "Yes")
{$Content += "<td bgcolor=#ff4000 align=center>Yes</td>"}
else
{$Content += "<td bgcolor=#65ff00 align=center>No</td>"}
$Content += "<td bgcolor=#f5f5f5 align=center>$WUboot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$PakBoot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$PakWowBoot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$RenFileBoot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$PCnameBoot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$CBSBoot</td>"
$Content += "</tr>"
}

#Close HTML
$Content += "</table>"
$Content += "</body>"
$Content += "</html>"

#Send Email Report
#Send-MailMessage -From $smtpFrom -To $smtpTo -Subject $Subject -Body $Content -BodyAsHtml -Priority High -dno onSuccess, onFailure -SmtpServer $smtpServer
$Content | Out-File C:\LazyWinAdmin\PendingReboot.html

