# api: multitool
# version: 0.9
# title: Server owner
# description: Find OU and app owner
# type: inline
# category: beta
# x_category: network
# icon: info
# key: n7|serverinfo
# config: -
#
# ❏ supply servername in computername field
#
# ❏ users Get-ADComputer $server for basic details
#
# ❏ data/servers.*.csv
#    - e.g. servers.office1.csv + servers.our_webhoster.csv + servers.under_cupboard.csv
#    - meant as custom extra info lists
#    - export as CSV with e.g. one "Hostname" field
#
# ❏ You might have to adapt the OU= extraction (or remove, if the AD
#   structure yields nothing interesting otherwise)


Param(
    $server = (Read-Host "Machine"),
    $automat = 0
)


#-- AD/OU
$server = $server.toUpper()
try {
    $ad = Get-ADComputer $server -Prop *
    $dn = $ad.DistinguishedName
    if ($dn -match "OU=(?!Server|Computer)(\w+)") {
        $dn = $matches[1]
    }
}
catch {
    Write-Host -f Yellow -b Red " ✘ No AD entry for $server"
    $ad = @{ DNSHostName=$server; OperatingSystem="OS=n/a"; SamAccountName="SAM=n/a" }
    $dn = "OU=n/a (shadowed/routed-only?)"
}
$site = "unknown site"
try {
    if ($x = (& nltest /server:$machine /dsgetsite)) {
        $site = $x[0]
    }
}
catch {}
try {
    $ip = [System.Net.Dns]::GetHostAddresses($machine)
}
catch {
    $ip = "not revolved"
}
Write-Host -f White  " ❏ $($ad.DNSHostName)"
Write-Host -f Yellow -b Red " ➜ $dn"
Write-Host -f Yellow " ⤷ $site"
Write-Host -f Yellow " ❏ IP=$ip"
Write-Host -f Yellow " ❏ $($ad.OperatingSystem)"
Write-Host -f Yellow " ❏ $($ad.SamAccountName)"


#-- Ping
if (Test-connection -quiet $server) {
    Write-Host -b Green " ✔ online"
}
else {
    Write-Host -b Red " ✘ not online"
}


#-- CSV infos (app owner)
ForEach ($fn in (dir "./data/servers.*.csv")) {
    if ($csv = Import-Csv $fn) {
        $r = ($csv | ? { $_.Hostname -eq $server -or $_."Virtual Machine" -eq $server })
    }
    if ($r) {
        Write-Host ""
        Write-Host -f Cyan " ❏ $($fn.Name)"
        ($r | FL | Out-String).trim() | Write-Host
    }
}

