## Looks for Old PC's by checking the ModificationDate of the Computer object. Computer passwords
## change every 30 days.

$date = Read-Host "Enter Date (i.e 1.15.2010)"
$ExecutionContext.InvokeCommand.ExpandString($date)
$file = "c:\Scripts\POSH\OLD_PC.txt"
Write-Host	"Creating Folder C:\Users\jbattista\Documents\OLD_PCs" -ForegroundColor "green"
if(!(Test-Path C:\Users\jbattista\Documents\OLD_PCs -PathType Container)){New-Item C:\Users\jbattista\Documents\OLD_PCs -type directory}
get-qadcomputer -SizeLimit 0 | Where-Object {$_.modificationdate -lt $date} | Select-Object name | export-csv $file -NoTypeInformation
(gc $file) -replace('"','')| out-file $file
Write-host `
	"Pinging List of Dead Computers and Writing to C:\Users\jbattista\Documents\OLD_PCs\PingOldPCs.csv `n WAIT" -foregroundcolor "green"
gc $file | ForEach-Object {GWMI win32_pingstatus -filter ("address='"+ $_ +"'") -computername.} | Select-Object address,responsetime,statuscode | export-csv C:\Users\jbattista\Documents\OLD_PCs\PingOldPCs.csv -NoTypeInformation
Write-Host "DONE. If a PC has a 0 then it responded to Ping" -ForegroundColor "green"