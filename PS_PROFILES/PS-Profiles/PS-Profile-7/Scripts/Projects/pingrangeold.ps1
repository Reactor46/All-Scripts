$n = Read-Host "What is the Network Address?(ie 192.168.1.)"
$i = Read-Host "What is the first address?"
$j = Read-Host "What is the last address?(ie 254)"
$m = [int]$i
do {
ping $n$m; $m++
}
while ($m -le $j)