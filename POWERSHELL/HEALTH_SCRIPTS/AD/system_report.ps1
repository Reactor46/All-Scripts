# system_report - (C) 2015 Patrick Lambert - http://dendory.net
#
# This should be run on the Task Scheduler every hour as a local administrator. 
# Comment out any section you don't want. The only resource intensive one seems to be Windows Updates.
# The AntiVirus/AntiMalware sections only work on workstations, they will return empty strings on Windows Server.
# Look near the bottom of the script to configure alarm report emails / notifications.
#
# The report will be saved in this file:
param([string]$File = ".\system_report.html")

"<html>
<head>
<link rel=stylesheet href=.\bootstrap.css>
<title>$env:COMPUTERNAME</title>
<style>
html, body
{
font-family: tahoma, arial;
font-size: 14px;
margin: 5px;
padding: 0px;
color: #2E1E2E;
}
table
{
border-collapse: collapse;
    width: 100%;
}
th
{
border: 1px solid #8899AA;
padding: 3px 7px 2px 7px;
font-size: 1.1em;
text-align: left;
padding-top: 5px;
padding-bottom: 4px;
background-color: #AABBCC;
color: #ffffff;
}
td
{
border: 1px solid #8899AA;
padding: 3px 7px 2px 7px;
    overflow: hidden;
}
h2
{
    text-align: center;
font-size: 22px;
    text-shadow: 1px 1px 1px rgba(150, 150, 150, 0.5);
}
h1
{
    margin-top: 20px;
    text-align: center;
font-size: 25px;
    text-shadow: 1px 1px 1px rgba(150, 150, 150, 0.5);
}
pre {
    white-space: pre-wrap;
    white-space: -moz-pre-wrap;
    white-space: -pre-wrap;
    white-space: -o-pre-wrap;
    word-wrap: break-word;
}
#sysinfo
{
    width: 49% !important;
    float: left;
    margin-bottom: 0px;
}
#action
{
    width: 49% !important;
    float: right;
}
#menu 
{
    position: fixed;
    right: 0;
    left: 0;
    top: 0;
    width: 100%;
    height: 25px;
    background: #AABBCC;
    color: #FFFFFF;
    text-align: center;
    overflow: hidden;
}
#menu a
{
    color: #FFFFFF;
    font-weight: bold;
}
@media screen and (max-width: 1010px)
{
    #sysinfo
    {
        float: none;
        margin-bottom: 20px;
        width: 100% !important;
    }
    #action
    {
        width: 100% !important;
        float: none;
    }
}
</style>
</head>
<body>
<div id='menu'>
<a href=#sysinfo>System</a> | <a href=#disks>Disks</a> | <a href=#network>Network</a> | <a href=#processes>Processes</a> | <a href=#services>Services</a>
</div>
<a name='sysinfo'></a><h1>$env:COMPUTERNAME System Report</h1>
" > $File

Write-Output "Fetching data:"
Write-Output "* Processor"
$processor = Get-WmiObject win32_processor
Write-Output "* System"
$sysinfo = Get-WmiObject win32_computersystem
Write-Output "* BIOS"
$bios = Get-WmiObject -Class win32_bios
Write-Output "* Operating System"
$os = Get-WmiObject win32_operatingsystem
Write-Output "* Users"
$users = Get-WmiObject win32_systemusers

"<table id='sysinfo'><tr><th colspan=2>System Information</th></tr>" >> $File
"<tr><td>Computer Name</td><td>" + $sysinfo.Name + "</td></tr>" >> $File
"<tr><td>Computer Type</td><td>" + $sysinfo.SystemType + "</td></tr>" >> $File
"<tr><td>Computer Manufacturer</td><td>" + $sysinfo.Manufacturer + "</td></tr>" >> $File
"<tr><td>Computer Model</td><td>" + $sysinfo.Model + "</td></tr>" >> $File
"<tr><td>CPU Information</td><td>" + $processor.Name + "</td></tr>" >> $File
"<tr><td>Installed RAM</td><td>" + [math]::Round($sysinfo.TotalPhysicalMemory / 1000000000) + " GB</td></tr>" >> $File
"<tr><td>BIOS Manufacturer</td><td>" + $bios.Manufacturer + "</td></tr>" >> $File
"<tr><td>BIOS Name</td><td>" + $bios.Name + "</td></tr>" >> $File
"<tr><td>BIOS Serial</td><td>" + $bios.SerialNumber + "</td></tr>" >> $File
"<tr><td>Hostname</td><td>" + $sysinfo.DNSHostName + "</td></tr>" >> $File
"<tr><td>Domain</td><td>" + $sysinfo.Domain + "</td></tr>" >> $File
"<tr><td>Operating System</td><td>" + $os.Caption + " (" + $os.OSArchitecture + ")</td></tr>" >> $File
"<tr><td>Build Number</td><td>" + (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion" |Select -ExpandProperty CurrentBuildNumber) + "</td></tr>" >> $File
"<tr><td>Product ID</td><td>" + (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion" |Select -ExpandProperty ProductId) + "</td></tr>" >> $File
"<tr><td>Local Users</td><td>" >> $File
ForEach ($u in $users) { $u.PartComponent -match ".*Name=(?<username>.*),.*Domain=(?<domain>.*).*" | Out-Null; $matches.username >> $File; " " >> $File }
"</td></tr>" >> $File
"</table>" >> $File

Write-Output "* Action Center"
$as = Get-WmiObject -Namespace "root\SecurityCenter2" -Class AntiSpywareProduct
$av = Get-WmiObject -Namespace "root\SecurityCenter2" -Class AntiVirusProduct
$fw_std = Get-ItemProperty "HKLM:System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile" | select -ExpandProperty EnableFirewall
$fw_dmn = Get-ItemProperty "HKLM:System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile" | select -ExpandProperty EnableFirewall
$fw_pub = Get-ItemProperty "HKLM:System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\PublicProfile" | select -ExpandProperty EnableFirewall
Write-Output "* Windows Updates"
$lastupd = Get-HotFix | Where-Object {$_.InstalledOn} | Sort-Object -Property InstalledOn | Select -Last 1 | Select -ExpandProperty InstalledOn
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$wu = $UpdateSearcher.Search("IsInstalled=0")
Write-Output "* System Load"
$cpuload = Get-Counter -Counter "\Processor(*)\% Processor Time" | Select -ExpandProperty CounterSamples | Select -ExpandProperty CookedValue | Measure-Object -Average | Select -ExpandProperty Average
$freemem = Get-Counter -Counter "\Memory\Available MBytes" | Select -ExpandProperty CounterSamples | Select -ExpandProperty CookedValue
$freemem = $freemem / 1000
$netload = [math]::round(((Get-Counter -Counter "\Network Interface(*)\Bytes Total/sec" -SampleInterval 1 -MaxSamples 3 | Select -ExpandProperty CounterSamples | Select -ExpandProperty CookedValue |Measure -Maximum |Select -ExpandProperty Maximum) / 1000),1)

"<table id='action'><tr><th colspan=2>Action Center</th></tr>" >> $File
"<tr><td>Anti-Virus Software</td><td>" + $av.displayName + " " + $av.VersionNumber + "</td></tr>" >> $File
"<tr><td>Anti-Spyware Software</td><td>" + $as.displayName + " " + $as.VersionNumber + "</td></tr>" >> $File
"<tr><td>Firewall Status</td><td>Domain: " + (&{If($fw_dmn -eq 1) {"On"} Else {"<font color=red>Off</font>"}}) + ", Private: " + (&{If($fw_std -eq 1) {"On"} Else {"<font color=red>Off</font>"}}) + ", Public: " + (&{If($fw_pub -eq 1) {"On"} Else {"<font color=red>Off</font>"}}) + "</td></tr>" >> $File
"<tr><td>Processor Load</td><td>" + (&{If($cpuload -lt 80) {[math]::Round($cpuload,2)} Else {"<font color=red>"+[math]::Round($cpuload,2)+"</font>"}}) + "%</td></tr>" >> $File
"<tr><td>Network Load</td><td>" + $netload + " KBytes/s</td></tr>" >> $File
"<tr><td>Free Memory</td><td>" + (&{If($freemem -gt 0.4) {"$freemem GB"} Else {"<font color=red>$freemem GB</font>"}}) + "</td></tr>" >> $File
"<tr><td>Last Boot</td><td>" + $os.ConvertToDateTime($os.LastBootUpTime) + " (" + (&{If($sysinfo.BootupState -eq "Normal boot") {$sysinfo.BootupState} Else {"<font color=red>"+$sysinfo.BootupState+"</font>"}}) + ")</td></tr>" >> $File
"<tr><td>Last Windows Update</td><td>" + $lastupd + (&{If(Get-ChildItem "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing" | Where {$_.PSChildName -eq "RebootPending"}) { " <font color=red>(Reboot pending)</font>" }}) + "</td></tr>" >> $File
"<tr><td>Available Critical Updates</td><td>" >> $File
$criticals = $wu.updates | where { $_.MsrcSeverity -eq "Critical" }
ForEach($critical in $criticals)
{
    "<font color=red>" >> $File
    $critical | Select -ExpandProperty Title >> $File
    "</font><br>" >> $File 
}
"</td></tr>" >> $File
Write-Host "* Event log"
$events = Get-EventLog Security -EntryType FailureAudit -After (Get-Date).AddHours(-1)
if($events)
{
    ForEach($event in $events) 
    {
        $id = $event.InstanceID
        $msg = $event.Message
      	$tim = $event.TimeGenerated
        "<tr><td>Event Audit Failure ($id)</td><td><font color=red><pre>$msg</pre>Time Generated: $tim</font></td></tr>" >> $File
    }
}
"</table><div style='clear:both'></div>" >> $File

"<a name='disks'></a><h2>Disk Space</h2>" >> $File
Write-Output "* Disks"
$disks = Get-WmiObject -Class win32_logicaldisk

"<table><tr><th>Drive</th><th>Type</th><th>Size</th><th>Free Space</th></tr>" >> $File
ForEach($d in $disks)
{
    $drive = $d.Name
    $type = $d.Description
    $size = [math]::Round($d.Size / 1000000000,1)
    $freespace = [math]::Round($d.FreeSpace / 1000000000,1)
    If($freespace -le 1 -And $freespace -ne 0) { "<tr><td>$drive</td><td>$type</td><td>$size GB</td><td><font color=red>$freespace GB</font></td></tr>" >> $File }
    Else { "<tr><td>$drive</td><td>$type</td><td>$size GB</td><td>$freespace GB</td></tr>" >> $File }
}
"</table>" >> $File

"<a name='network'></a><h2>Network Addresses</h2>" >> $File
Write-Output "* Network"
Get-WmiObject -Class 'Win32_NetworkAdapterConfiguration' -Filter 'IPEnabled = True' | Select @{Name='Interface';Expression={$_.Description}},@{Name='IP Addresses';Expression={$_.IPAddress}} | ConvertTo-Html -Fragment >> $File

"<a name='processes'></a><h2>Running Processes</h2>" >> $File
Write-Output "* Processes"
Get-WmiObject -Class win32_process | Sort -Property WorkingSetSize -Descending | Select @{Name='ID';Expression={$_.ProcessId}},@{Name='Name';Expression={$_.ProcessName}},@{Name='Path';Expression={$_.CommandLine}},@{Name='Memory Usage (MB)';Expression={[math]::Round($_.WorkingSetSize / 1000000, 3)}} | ConvertTo-Html -Fragment >> $File

"<a name='services'></a><h2>Running Services</h2>" >> $File
Write-Output "* Services"
Get-WmiObject -Class win32_service -Filter 'Started=True' | Sort -Property DisplayName | Select @{Name='Name';Expression={$_.DisplayName}},@{Name='Mode';Expression={$_.StartMode}},@{Name='Path';Expression={$_.PathName}},Description | ConvertTo-Html -Fragment >> $File

$date = Get-Date
"<p><i>Report produced: $date</i></p>" >> $File

if((Get-Content $File | Select-String -Pattern "color=red"))
{
    Write-Output "*** Alarms were raised!"
    # Uncomment this out to send an email:
    #Send-MailMessage -From "noreply@example.com" -To "somewhere@example.com" -Subject "System Report" -Body (Get-Content $File) -BodyAsHtml -SmtpServer "localhost"
    # Uncomment this out to use pushbullet to send a notification:
    #pushbullet.exe -apikey APIKEY -title "System Report" -link "http://link/to/this/report.html"
}

Write-Output "Done! Report at: $File"
