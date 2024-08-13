# DeleteOldIISLogFiles.ps1 
# Author: Ronni Pedersen (www.ronnipedersen.com)
# Last update: 2014.06.24
# Version: 1.0
# 
# This script scans all *.log files in a folder (including subfolders), for files that haven't be used for x numbers of days.
# This Script was created for remediation use with SCCM 2012 Compliance Settings, but can be used outside SCCM.
# Example use: Delete all IIS log files that is older that 7 days.
#
# Credits and thanks goes to these guys:
# This script is inspired from Thomas Kurth, found at (http://netecm.netree.ch/blog/Lists/Posts/Post.aspx?ID=79)
#
# Feedback are always welcome, please feedback at www.ronnipedersen.com.

########################
# Variables to set

$Now = Get-Date
$Days = "7"
$TargetFolder="C:\inetpub\logs\LogFiles"
$Extension = "*.log"
$LastWrite = $Now.AddDays(-$Days)

##########################

if(Test-Path $TargetFolder ){
    $files = get-childitem -Path $TargetFolder -Include $Extension -Recurse | Where {$_.lastwritetime -lt "$LastWrite"} 

    foreach ($File in $Files) {
        write-host "Deleting File $File" -ForegroundColor Red
        Remove-Item $File.FullName | out-null
    }
}
Else{
    Write-host "The Target folder does not exists" -ForegroundColor Green
}