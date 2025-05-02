cd \
cd .\Scripts
$Servers = Get-Content exchange.txt
Foreach ($Server in $Servers)
{
Write-Host -backgroundcolor yellow **************$Server*********************
# Get-EventLog -ComputerName $Server -LogName Application -After (Get-Date).date | Where-Object {$_.EventID -eq 10024} | Get-Unique
Get-EventLog -ComputerName $Server -LogName Application  | Where-Object {$_.EventID -eq 10024} | Get-Unique
}