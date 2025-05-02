$servers = Get-Content C:\LazyWinAdmin\IIS\IIS-Reports\Web.txt 
#Enumerate the server list from the text file
foreach ($server in $servers) {

Invoke-Command -ComputerName $server -ScriptBlock {Get-ChildItem -Path Cert:\LocalMachine\My | Select -Property * } }