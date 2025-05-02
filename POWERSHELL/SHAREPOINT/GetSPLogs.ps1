Set-sploglevel -traceseverity verboseex
 
New-SPlogfile
 
$starttime = Get-Date
 
# reproduce the issue or load fiddler 
 
$endtime = Get-Date
 
Clear-SPLogLevel 
 
New-Splogfile
 
Merge-SPLogFile -Path C:\SPLogs.log -StartTime $starttime -EndTime $endtime