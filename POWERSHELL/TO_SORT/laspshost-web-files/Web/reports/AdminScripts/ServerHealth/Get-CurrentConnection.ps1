function Get-CurrentConnection($Site) {
Get-Counter "\web service($Site)\current connections","\web service($Site)\Bytes Received/sec","\web service($Site)\Bytes Sent/sec" -ComputerName $env:COMPUTERNAME
}

$IISsites = Get-Website | Select Name
$IISsites = $IISsites.Name
$CurrentConnection = @()
foreach ($site in $IISsites)
{
Write-Host $Site
$ConnCount = New-Object psobject | Get-CurrentConnection -Site $Site
$CurrentConnection += $ConnCount
}
$CurrentConnection | out-gridview
