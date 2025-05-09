﻿<#
.SYNOPSIS
    Gets the Windows Volumes on the target server
	
.DESCRIPTION
   Writes the windows Volumes out to the "01 - Server Storage" folder
   One file for all volumes
   This is to know what drive and mount-points the server has
   
.EXAMPLE
    01_Server_Storage.ps1 localhost
	
.EXAMPLE
    01_Server_Storage.ps1 server01 sa password

.Inputs
    ServerName, [SQLUser], [SQLPassword]

.Outputs
	HTML File
	
.NOTES

	
.LINK
	https://github.com/gwalkey
	
#>

[CmdletBinding()]
Param(
  [string]$SQLInstance='localhost',
  [string]$myuser,
  [string]$mypass
)

# Load Common Modules and .NET Assemblies
try
{
    Import-Module ".\SQLTranscriptase.psm1" -ErrorAction Stop
}
catch
{
    Throw('SQLTranscriptase.psm1 not found')
}

LoadSQLSMO

# Init
Set-StrictMode -Version latest;
[string]$BaseFolder = (Get-Item -Path ".\" -Verbose).FullName
Write-Host  -f Yellow -b Black "01 - Server Storage"
Write-Output("Server: [{0}]" -f $SQLInstance)

# Output folder
$fullfolderPath = "$BaseFolder\$sqlinstance\01 - Server Storage\"
if(!(test-path -path $fullfolderPath))
{
    mkdir $fullfolderPath | Out-Null
}

# Split out servername only from named instance
$WinServer = ($SQLInstance -split {$_ -eq "," -or $_ -eq "\"})[0]

# Credit: https://sqlscope.wordpress.com/2012/05/05/
$VolumeTotalGB = @{Name="VolumeTotalGB";Expression={[Math]::Round(($_.Capacity/1GB),2)}}
$VolumeUsedGB =  @{Name="VolumeUsedGB";Expression={[Math]::Round((($_.Capacity - $_.FreeSpace)/1GB),2)}}
$VolumeFreeGB =  @{Name="VolumeFreeGB";Expression={[Math]::Round(($_.FreeSpace/1GB),2)}}

# Let WMI errors be trapped
try
{
	$VolumeArray = Get-WmiObject -Computer $WinServer Win32_Volume -ErrorAction stop| sort-object name 
    if ($?)
    {
        Write-Output "Good WMI Connection"
    }
    else
    {
        $fullfolderpath = "$BaseFolder\$SQLInstance\"
        if(!(test-path -path $fullfolderPath))
        {
            mkdir $fullfolderPath | Out-Null
        }
        echo null > "$fullfolderpath\01 - Server Storage - WMI Could not connect.txt"
        Write-Output "WMI Could not connect"
        Set-Location $BaseFolder
        Throw("WMI Could not connect")
  
    }
}
catch
{
    $fullfolderpath = "$BaseFolder\$SQLInstance\"
    if(!(test-path -path $fullfolderPath))
    {
         mkdir $fullfolderPath | Out-Null
    }
    echo null > "$fullfolderpath\01 - Server Storage - WMI Could not connect.txt"
    Write-Output "WMI Could not connect"        
    Set-Location $BaseFolder
    exit
}

# HTML CSS
$head = "<style type='text/css'>"
$head+="
table
    {
        Margin: 0px 0px 0px 4px;
        Border: 1px solid rgb(190, 190, 190);
        Font-Family: Tahoma;
        Font-Size: 9pt;
        Background-Color: rgb(252, 252, 252);
    }
tr:hover td
    {
        Background-Color: rgb(150, 150, 220);
        Color: rgb(255, 255, 255);
    }
tr:nth-child(even)
    {
        Background-Color: rgb(242, 242, 242);
    }
th
    {
        Text-Align: Left;
        Color: rgb(150, 150, 220);
        Padding: 1px 4px 1px 4px;
    }
td
    {
        Vertical-Align: Top;
        Padding: 1px 4px 1px 4px;
    }
"
$head+="</style>"

$RunTime = Get-date
Write-Output('{0} Volumes found' -f @($VolumeArray).Count)
$myoutputfile4 = $FullFolderPath+"\Server_Storage_Volumes.html"
$myHtml1 = $VolumeArray | select Name, Label, FileSystem, DriveType, $VolumeTotalGB, $VolumeUsedGB, $VolumeFreeGB, BootVolume, DriveLetter, BlockSize | `
ConvertTo-Html -Fragment -as table -PreContent "<h1>Server: $SqlInstance</H1><H2>Storage Volumes</h2>"
Convertto-Html -head $head -Body "$myHtml1" -Title "Storage Volumes"  -PostContent "<h3>Ran on : $RunTime</h3>" | Set-Content -Path $myoutputfile4

# Return To Base
set-location $BaseFolder
