$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://LASEXDB01.Contoso.CORP/PowerShell/ -Authentication Kerberos
Import-PSSession -Session $Session -AllowClobber

$Servers = Get-Content C:\LazyWinAdmin\Exchange\Scripts\Exchange.txt
Foreach ($Server in $Servers)
{
Write-Host -backgroundcolor yellow **************$Server*********************
# Get-EventLog -ComputerName $Server -LogName Application -After (Get-Date).date | Where-Object {$_.EventID -eq 10024} | Get-Unique
Get-EventLog -ComputerName $Server -LogName Application  | Where-Object {$_.EventID -eq 10024} | Get-Unique
}


$Servers = Get-Content C:\LazyWinAdmin\Exchange\Scripts\Exchange.txt
Foreach ($Server in $Servers)
{
Write-Host -backgroundcolor yellow **************$Server*********************
# Get-EventLog -ComputerName $Server -LogName Application -After (Get-Date).date | Where-Object {$_.EventID -eq 10024} | Get-Unique
Get-EventLog -ComputerName $Server -LogName Application  | Where-Object {$_.EventID -eq 10023} | Get-Unique
}