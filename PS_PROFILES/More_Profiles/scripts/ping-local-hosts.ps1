﻿<#
.SYNOPSIS
        Pings local hosts
.DESCRIPTION
        This PowerShell script pings the computers in the local network and lists which one are up.
.EXAMPLE
        PS> ./ping-local-hosts.ps1
	✅ Up: Hippo Jenkins01 Jenkins02 Rocket Vega 
.LINK
        https://github.com/fleschutz/PowerShell
.NOTES
        Author: Markus Fleschutz | License: CC0
#>

try {
	Write-Progress "Sending pings to the local hosts..."
	[string]$hosts = "Amnesiac,ArchLinux,Berlin,Boston,Brother,Canon,Castor,Cisco,EchoDot,Epson,Fedora,Fireball,Firewall,fritz.box,GasSensor,Gateway,Hippo,HomeManager,Io,iPhone,Jarvis,Jenkins01,Jenkins02,LA,Laptop,Jupiter,Mars,Mercury,Miami,Mobile,NY,OctoPi,Paris,Pixel-6a,Pluto,Printer,Proxy,R2D2,Raspberry,Rocket,Rome,Router,Server,Shelly1,SmartPhone,SmartWatch,Soundbar,Sunnyboy,Surface,Switch,Tablet,Tolino,TV,Ubuntu,Vega,Venus,XRX,Zeus" # sorted alphabetically
	$hostsArray = $hosts.Split(",")
	$count = $hostsArray.Count

	[int]$timeout = 600 # ms
        $queue = [System.Collections.Queue]::new()
	foreach($hostname in $hostsArray) {
		$ping = [System.Net.Networkinformation.Ping]::new()
		$obj = @{ Host = $hostname; Ping = $ping; Async = $ping.SendPingAsync($hostname, $timeout) }
 		$queue.Enqueue($obj)
        }

	[string]$result = ""
	while ($queue.Count -gt 0) {
		$obj = $queue.Dequeue()
		try {
                	if ($obj.Async.Wait($timeout) -eq $true) {
				if ($obj.Async.Result.Status -ne "TimedOut") {
					$result += "$($obj.Host) "
				}
				continue
			}
		} catch {
			if ($obj.Async.IsCompleted -eq $true) {	continue }
		}
		$queue.Enqueue($obj)
	}
	Write-Progress -completed "Done."
	Write-Host "✅ Up: $($result)"
	exit 0 # success
} catch {
        "⚠️ Error in line $($_.InvocationInfo.ScriptLineNumber): $($Error[0])"
        exit 1
}
