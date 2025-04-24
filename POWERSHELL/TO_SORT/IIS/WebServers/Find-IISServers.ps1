$servers = (Get-Content C:\LazyWinAdmin\AntiVirus\Symantec\ALL-SERVERS.txt)
 
foreach($vm in $servers){
$iis = get-wmiobject Win32_Service -ComputerName $vm -Filter "name='IISADMIN'"
 
if($iis.State -eq "Running")
{Write-Output "IIS is running on $vm" | Out-File 'C:\LazyWinAdmin\WebServers\IIS-Running.log' -append}
 
else
{Write-Output "IIS is not running on $vm"| Out-File 'C:\LazyWinAdmin\WebServers\IIS-NotInstalled.log' -append}
}