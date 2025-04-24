<# 
.SYNOPSIS 
Imports all-user WLAN profiles based on Xml-files in the specified directory 
.DESCRIPTION 
Imports all-user WLAN profiles based on Xml-files in the specified directory using netsh.exe 
.PARAMETER XmlDirectory 
Directory to import Xml configuration-files from 
.EXAMPLE 
Import-WLAN -XmlDirectory c:\temp\wlan 
#> 
 
[CmdletBinding()] 
    param ( 
        [parameter(Mandatory=$true)] 
        [string]$XmlDirectory 
        ) 
 
#Import all WLAN Xml-files from specified directory 
Get-ChildItem $XmlDirectory | Where-Object {$_.extension -eq ".xml"} | ForEach-Object {netsh wlan add profile filename=($XmlDirectory+"\"+$_.name)} 