################################################################################################
# Powershell Backup CA Authority
# Author: Martijn Kamminag
# https://www.isee2it.nl
# Date: 13 januari 2018
# Version: 1.0
#
# THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
# OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
# FITNESS FOR A PARTICULAR PURPOSE.
#
# This sample is not supported under any Microsoft standard support program or service. 
# The script is provided AS IS without warranty of any kind. Microsoft further disclaims all
# implied warranties including, without limitation, any implied warranties of merchantability
# or of fitness for a particular purpose. The entire risk arising out of the use or performance
# of the sample and documentation remains with you. In no event shall Microsoft, its authors,
# or anyone else involved in the creation, production, or delivery of the script be liable for 
# any damages whatsoever (including, without limitation, damages for loss of business profits, 
# business interruption, loss of business information, or other pecuniary loss) arising out of 
# the use of or inability to use the sample or documentation, even if Microsoft has been advised 
# of the possibility of such damages.
################################################################################################
# Run with Highest Privileges in Scheduled Tasks - aka RunAsAdmin in Powershell
################################################################################################

# Setting variables
$source = Get-ChildItem C:\Windows\System32\CertLog\ -Recurse -Exlude tmp.edb | Where-Object { $_.Extension -Match "edb" } | Select-Object Name
$DBCA = $source.Name
$backupdestination = "D:\CA-Backup"
$curr_date = Get-Date
$max_retention = "-8"
$deletefoldersolderthen = $curr_date.AddDays($max_retention)
$folderdate = (Get-Date).ToString("dd-MM-yy")
$computer = $env:computername
$passwordforbackupCA = "PASSWORD"
$smtpServer = "IP/FQDN"
$sendfrom = "your@email.com"
$sendto = "your@email.com"

# Folder Check, Create if not exist
If (!(Test-Path $backupdestination )) { New-Item $backupdestination -Type Directory } else { write-host allready exists }
takeown /f $backupdestination /r /d Y
If (!(Test-Path $backupdestination\$folderdate)) { New-Item $backupdestination\$folderdate -Type Directory } else { write-host allready exists }
If (!(Test-Path $backupdestination\$folderdate\Policy)) { New-Item $backupdestination\$folderdate\Policy -Type Directory } else { write-host allready exists }
If (!(Test-Path $backupdestination\$folderdate\Register)) { New-Item $backupdestination\$folderdate\Register -Type Directory } else { write-host allready exists }
If (!(Test-Path $backupdestination\$folderdate\Backup)) { New-Item $backupdestination\$folderdate\Backup -Type Directory } else { write-host allready exists }
If (!(Test-Path $backupdestination\$folderdate\Templates)) { New-Item $backupdestination\$folderdate\Templates -Type Directory } else { write-host allready exists }
takeown /f $backupdestination\$folderdate\Templates /r /d Y
# Backup the Registry
reg.exe export HKLM\System\CurrentControlSet\Services\certsvc $backupdestination\$folderdate\Register\certsvc.reg

# Backup the CAPolicy.inf
robocopy C:\Windows D:\CA-Backup\$folderdate\Policy CAPolicy.inf

# Backup the Database 
certutil -backup -p $passwordforbackupCA -f -gmt -seconds -v $backupdestination\$folderdate\Backup
certutil -backupKey -p $passwordforbackupCA -f -gmt -seconds -v $backupdestination\$folderdate\BackupKey
certutil.exe -v -catemplates > $backupdestination\$folderdate\Templates\Templates.txt

#Count the number of Backup Databases
$CurrentBackupsCA = Get-ChildItem $backupdestination -Recurse | Where-Object { $_.Extension -Match "edb" } | Measure-Object | %{$_.Count}

# Prepare for mail in case of success/failure
if (Test-Path -Path "$backupdestination\$folderdate\Backup\DataBase\$DBCA") { 
$subject="$computer Backup CA has completed sucessfully" 
$body="$computer Backup CA has completed sucessfully. Total of $CurrentBackupsCA Backups" 
} 
else 
{ 
$subject="$computer Backup Issuing CA has failed" 
$body="$computer the $backupdestination\$folderdate\Backup\DataBase\$DBCA File could not be found, Please check your server. Total backups found: $CurrentBackupsCA" 
} 
# Send the mail
send-mailmessage -from $sendfrom -to $sendto -subject $subject -BodyAsHTML -body $body -priority high -smtpServer $smtpServer 
# Delete based on retention variables
Get-ChildItem $backupdestination -Recurse | Where-Object { $_.CreationTime -lt $deletefoldersolderthen } | Remove-Item -Recurse -Force -Confirm:$False
