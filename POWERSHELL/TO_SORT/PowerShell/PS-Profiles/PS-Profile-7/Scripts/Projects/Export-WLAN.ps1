 
<# 
.SYNOPSIS 
Exports all-user WLAN profiles 
.DESCRIPTION 
Exports all-user WLAN profiles to Xml-files to the specified directory using netsh.exe 
.PARAMETER XmlDirectory 
Directory to export Xml configuration-files to 
.EXAMPLE 
Export-WLAN -XmlDirectory c:\temp\wlan 
#> 
 
[CmdletBinding()] 
    param ( 
        [parameter(Mandatory=$true)] 
        [string]$XmlDirectory 
        ) 
 
#Export all WLAN profiles to specified directory 
$wlans = netsh wlan show profiles | Select-String -Pattern "All User Profile" | Foreach-Object {$_.ToString()} 
$exportdata = $wlans | Foreach-Object {$_.Replace("    All User Profile     : ",$null)} 
$exportdata | ForEach-Object {netsh wlan export profile $_ $XmlDirectory key=clear} 
