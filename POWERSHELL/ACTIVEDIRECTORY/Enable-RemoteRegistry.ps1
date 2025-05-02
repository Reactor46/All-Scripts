$computers = Get-ADComputer -LDAPFilter "(name=*DPC*)" -SearchBase "OU=Computers,OU=Collections,OU=Las_Vegas,DC=contoso,DC=com" -ErrorAction SilentlyContinue -Properties *
foreach ($computer in $computers) 
{ 
if (Test-Connection -count 1 -computer $computer.Name -quiet){ 
Write-Host "Updating system" $computer.Name "....." -ForegroundColor Green 
Set-Service –Name remoteregistry –Computer $computer.Name -StartupType Automatic 
Get-Service remoteregistry -ComputerName $computer.Name | start-service 
} 
else 
{ 
Write-Host "System Offline " $computer.Name "....." -ForegroundColor Red 
echo $computer.Name >> C:\LazyWinAdmin\offlineRemoteRegStartup.txt} 
}