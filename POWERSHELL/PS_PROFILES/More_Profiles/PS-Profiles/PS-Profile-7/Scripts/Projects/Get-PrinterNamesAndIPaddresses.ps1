	
	$ipaddress = @{Name="IP Address";Expression={$portname = $_.portname;(Get-WmiObject win32_tcpipprinterport -Filter Name='"'$portname'"' -ComputerName bhmckdc).hostaddress}}
	get-wmiobject win32_printer -computername bhmckdc -Property name, Portname | select name, $ipaddress | export-csv $desk\BHPrinters.csv
	