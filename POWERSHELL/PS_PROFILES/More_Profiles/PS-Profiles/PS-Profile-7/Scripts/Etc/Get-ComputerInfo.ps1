$Headings = "#ffcc5c"
$Content = "#ffeead"
$Background = "#96ceb4"
$TitleColor = "Black"
$FontColor = "Black"
$a = "
<title>Computer Info</title>
<style>
H1{color:$TitleColor; margin: auto;width: 60%;padding: 10px}
BODY{background-color:$Background;font-family: Arial;font-size: 20pt;Color: $FontColor}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:$Headings}
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:$Content}
div{margin: auto;width: 60%;padding: 10px;}
#host td:nth-child(1){ background-color:$Headings}
#vpn{}
</style>"

Function Make-IPTable
	{
	$ipinfo = get-wmiobject win32_networkadapterconfiguration | Where-Object { $_.ipaddress -ne $null } | select Description, IPaddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder, MacAddress
	$type = @{Name="Device";Expression={[string]$description = $_.description;(Get-WmiObject win32_NetworkAdapter -Property NetConnectionID  -Filter Name='"'$description'"').netconnectionid}}
	$iptable = $ipinfo | select $type, @{Name="IP Address";Expression={[string]$_.ipaddress}}, @{Name="Default Gateway";Expression={[string]$_.DefaultIPGateway}}
	[string]$IPTable = $IPTable | ConvertTo-Html
	$IPTable -match "<body>(.*?)</body>" | Out-Null
	$IPTable = $matches[1]
	Return $IPTable
	}

Function Make-HostTable
	{
	[string]$Hostname = ($env:COMPUTERNAME).tolower()
	[string]$Username = $env:USERNAME
	[string]$Boottime = (Get-CimInstance -ClassName win32_operatingsystem).lastbootuptime | Get-Date -Format "MMM d H:mm tt"
	$HostTable = "
<table>
<tr>
	<td><b>Computer Name</b></td>
	<td>$hostname</td>
</tr>
<tr>
	<td><b>User</b></td>
	<td>$username</td>
</tr>
<tr>
	<td><b>Last Boot</b></td>
	<td>$Boottime</td>
</tr>
</table>
"
	Return $HostTable
	}
	
Function Make-VPNTable
	{
	$nics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()
	$ppp = $nics | ? { $_.NetworkInterfaceType -eq "ppp" }
	if ( $ppp )
		{
		$VPNtable = $ppp | select @{Name="VPN Name";Expression="Name"}, @{Name="IP Address";Expression={($ppp.getipproperties()).unicastaddresses.Address}}
		[string]$VPNtable = $VPNtable | ConvertTo-Html
		$VPNTable -match "<body>(.*?)</body>" | Out-Null
		$VPNTable = $matches[1]
		}
	else
		{
		$VPNtable = ""
		}
	Return $VPNTable
	}
	
Function Generate-Page
	{
	$header = "" | ConvertTo-Html -Head $a | Out-String
	$header = $header.substring(0, $header.indexof("<body>"))
	$hosttable = Make-HostTable
	$iptable = Make-IPTable
	$VPNtable = Make-VPNTable
	$body = "
<body>
	<h1></h1>
	<div id='host'>
	$Hosttable
	</div>
	<div id='vpn'>
	$VPNTable
	</div>
	<div id='ip'>
	$IPTable
	</div>
</body>
"
	Return $header + $body 
	}
$htmlpath = Join-Path $env:temp compinfo.htm
if ( Test-Path $htmlpath ) { Remove-Item $htmlpath -Force }
Generate-Page | Out-File $htmlpath -Force
Invoke-Item $htmlpath
	

