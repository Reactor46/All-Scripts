#$command = "(get-wmiobject win32_networkadapterconfiguration | Where { $_.ipaddress -ne $null -and $_.Description -match 'Gigabit Network Connection' }).ipaddress"
#$lookupip = "powershell -command $command"
#$lookupip = "$(ip route get 8.8.8.8 | awk '{ print $NF; exit }')"

$name = "pihole"
$volume1 = "c:\users\wreeves\Dockers\pihole\:/etc/pihole/"
$volume2 = "c:\users\wreeves\Dockers\pihole\dnsmasq.d\:/etc/dnsmasq.d/"


docker run -d --name $name -p 53:53/tcp -p 53:53/udp -p 80:80 -v $volume1 -v $volume2 -e ServerIP=192.168.103.129 -e DNS1="192.168.10.110" -e DNS2="192.168.10.112" --restart=always diginc/pi-hole:debian
