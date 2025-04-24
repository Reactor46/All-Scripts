param($ip)

$mac = ((((arp -a | grep $ip) -replace '(^\s+|\s+$)', '' -replace '\s+',' ') -split ' ')[1]) -replace '-',':'
Send-Wol $mac