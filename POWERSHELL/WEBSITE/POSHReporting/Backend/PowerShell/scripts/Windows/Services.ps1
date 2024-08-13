$SharePointVersion = (Get-PSSnapin microsoft.sharepoint.powershell).Version.Major

$Services = Get-Service | Where-Object {$_.Displayname -like "*Sharepoint*"} | Select -ExpandProperty Name

if($SharePointVersion -eq 14)
{
    $ComputerName | Get-SPServices -Service $Services | Select "Server Name", "Display Name", "Name", "State", "Startup Type", "Status" | Group-ByStatus
}
elseif($SharePointVersion -eq 15)
{
    $ComputerName | Get-SPServices -Service $Services | Select "Server Name", "Display Name", "Name", "State", "Startup Type", "Status" | Group-ByStatus
}
elseif($SharePointVersion -eq 16)
{
    $ComputerName | Get-SPServices -Service $Services | Select "Server Name", "Display Name", "Name", "State", "Startup Type", "Status" | Group-ByStatus
} 
