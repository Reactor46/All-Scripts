GoGo-PSExch
$ExchangeServers = Get-ExchangeServer  | Sort-Object Name
ForEach  ($Server in $ExchangeServers)
{
Invoke-Command -ComputerName $Server.Name -ScriptBlock {Get-Command  Exsetup.exe | ForEach-Object {$_.FileversionInfo}} | Select PSComputername, FileVersion, ProductVersion
}