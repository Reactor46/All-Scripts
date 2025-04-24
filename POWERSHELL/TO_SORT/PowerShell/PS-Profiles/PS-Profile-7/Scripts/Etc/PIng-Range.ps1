param(
$Subnet,
$Start,
$End
)

$n = $Subnet
$s = [int]$start
$l = [int]$End
$sint = [int]$Start


do {
if (test-connection $n$sint -count 1 -buffersize 16 -erroraction stop -Quiet)
	{
	$ip = "$n$sint"
	try
		{
		$Hostname = [System.Net.Dns]::GetHostbyAddress($ip).hostname
		write-host "SUCCESSFUL: $ip[$hostname]" -ForegroundColor Green
		$reachable = $reachable + "$ip[$hostname]"
		}
	catch
		{
		write-host "SUCCESSFUL: $ip[HOSTNAME: UNRESOLVED]" -ForegroundColor Red
		$reachable = $reachable + "$ip[$hostname]" 
		}
	}
else
	{
	$ip = "$n$sint"
	Write-Host "UNREACHABLE $n$sint"
	$unreachable = $unreachable + "$ip"
	}
$sint++
}
while ($sint -le $l)
