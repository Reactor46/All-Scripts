﻿<#
.SYNOPSIS
    Gets the Windows SMB Shares on the target server
	
.DESCRIPTION
   Writes the SMB Shares out to the "01 - Server Shares" folder
   One file for all shares
   
.EXAMPLE
    Usage: 01_Server_Shares.ps1 `"SQLServerName`" ([`"Username`"] [`"Password`"] if DMZ machine)

.Inputs
    ServerName, [SQLUser], [SQLPassword]

.Outputs
	HTML Files
	
.NOTES

	
.LINK
	https://github.com/gwalkey
	
#>

[CmdletBinding()]
Param(
  [string]$SQLInstance="localhost",
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
Write-Host  -f Yellow -b Black "01 - Server Shares"
Write-Output("Server: [{0}]" -f $SQLInstance)

# Shares go here
$ShareArray = [System.Collections.ArrayList]@()

# WMI connects to the Windows Server Name, not the SQL Server Named Instance
$WinServer = ($SQLInstance -split {$_ -eq "," -or $_ -eq "\"})[0]

# Output folder
$fullfolderPath = "$BaseFolder\$sqlinstance\01 - Server Shares\"
if(!(test-path -path $fullfolderPath))
{
    mkdir $fullfolderPath | Out-Null
}

# Turn off default PS error handling - let them filter down from the WMI Call
$old_ErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

try
{

    $ShareArray = Get-WmiObject -Computer $WinServer -class Win32_Share | select Name, Path, Description | Where-Object -filterscript {$_.Name -ne "ADMIN$" -and $_.Name -ne "IPC$"} | sort-object name
    # $ShareArray | Out-GridView
    if ($?)
    {
        Write-Output "Good WMI Connection"
    }
    else
    {   
        echo null > "$fullfolderpath\01 - Server Shares - WMI Could not connect.txt"
        Write-Output "WMI Could not connect"
        Set-Location $BaseFolder
        exit
    }
}
catch
{
    $fullfolderpath = "$BaseFolder\$SQLInstance\"
    if(!(test-path -path $fullfolderPath))
    {
        mkdir $fullfolderPath | Out-Null
    }
    echo null > "$fullfolderpath\01 - Server Shares - WMI Could not connect.txt"
    Write-Output "WMI Could not connect"       
    Set-Location $BaseFolder
    exit
}


# Reset default PS error handler
$ErrorActionPreference = $old_ErrorActionPreference 

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


# Export It
$RunTime = Get-date

$myoutputfile4 = $FullFolderPath+"\Shares_Overview.html"
$myHtml1 = $ShareArray | select  Name, Path, Description | `
ConvertTo-Html -Fragment -as table -PreContent "<h1>Server: $SqlInstance</H1><H2>Shares Overview</h2>"
Convertto-Html -head $head -Body "$myHtml1" -Title "Shares Overview"  -PostContent "<h3>Ran on : $RunTime</h3>" | Set-Content -Path $myoutputfile4

# Loop Through Each Share, exporting NTFS and SMB permissions
Write-Output "Exporting NTFS/SMB Share Permissions..."


$PermPath = "$BaseFolder\$sqlinstance\01 - Server Shares\NTFS_Permissions\"
if(!(test-path -path $PermPath))
{
    mkdir $PermPath | Out-Null
}
$permpathfile = $PermPath + "NTFS_Permissions.txt"
"NTFS File Permissions for $Winserver shares`r" | out-file -FilePath $permpathfile -encoding ascii

$SMBPath = "$BaseFolder\$sqlinstance\01 - Server Shares\SMB_Permissions\"
if(!(test-path -path $SMBPath))
{
    mkdir $SMBPath | Out-Null
}
$SMBPathfile = $SMBPath + "SMB_Permissions.txt"
"SMB Share Permissions for $Winserver shares`r" | out-file -FilePath $SMBPathfile -encoding ascii

Function Get-NtfsRights($name,$path,$comp)
{
	$path = [regex]::Escape($path)
	$share = "\\$comp\$name"

    $old_ErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'

	$wmi = gwmi Win32_LogicalFileSecuritySetting -filter "path='$path'" -ComputerName $comp
	$wmi.GetSecurityDescriptor().Descriptor.DACL | where {$_.AccessMask -as [Security.AccessControl.FileSystemRights]} |select `
				@{name="Principal";Expression={"{0}\{1}" -f $_.Trustee.Domain,$_.Trustee.name}},
				@{name="Rights";Expression={[Security.AccessControl.FileSystemRights] $_.AccessMask }},
				@{name="AceFlags";Expression={[Security.AccessControl.AceFlags] $_.AceFlags }},
				@{name="AceType";Expression={[Security.AccessControl.AceType] $_.AceType }},
				@{name="ShareName";Expression={$share}}

    # Reset default PS error handler - for WMI error trapping
    $ErrorActionPreference = $old_ErrorActionPreference 
}

foreach($Share in $ShareArray)
{
    # Skip certain shares
    if ($Share.name -eq "print$") {continue}
    if ($Share.name -eq "FILESTREAM") {continue}
    if ($Share.name -eq "IPC$") {continue}
    if ($Share.name -eq "ADMIN$") {continue}
    
    # Skip shares with spaces in the path  - can you even connect to these?
    # $Share.path

    if ($Share.path.Contains(' '))
     {
        Write-Output ("--> Could not script out Share [{0}], with Path [{1}]" -f $share.name, $share.path)
        continue
     }

    # Get Security Descriptors on NTFS for the share
    $acl = Get-NtfsRights $Share.Name $Share.Path $WinServer

    # Enum
    foreach($accessRule in $acl)
    {
        Write-Output ("Share: {0}, Path: {1}, Identity: {2}, Rights: {3}" -f $accessRule.ShareName, $Share.path, $accessRule.Principal, $accessRule.Rights)
        Write-Output ("Share: {0}, Path: {1}, Identity: {2}, Rights: {3}" -f $accessRule.ShareName, $Share.path, $accessRule.Principal, $accessRule.Rights) | out-file -FilePath $permpathfile -append -encoding ascii
    }

    Write-Output ("`r`n") | out-file -FilePath $permpathfile -append -encoding ascii
   
    # Get Share SMB Perms
    $ShareName = $Share.Name
    $SMBShare = Get-WmiObject win32_LogicalShareSecuritySetting -Filter "name='$ShareName'" -ComputerName $WinServer
    if($SMBShare)
    {
        $obj = @()
        $ACLS = $SMBShare.GetSecurityDescriptor().Descriptor.DACL
        foreach($ACL in $ACLS)
        {
            $User = $ACL.Trustee.Name
            if(!($user)){$user = $ACL.Trustee.SID}
            $Domain = $ACL.Trustee.Domain
            switch($ACL.AccessMask)
            {
                1179785        {$Perm = "Read"}
                1180063        {$Perm = "Read, Write"}
                1179817        {$Perm = "ReadAndExecute"}
                -1610612736    {$Perm = "ReadAndExecuteExtended"}
                1245631        {$Perm = "ReadAndExecute, Modify, Write"}
                1180095        {$Perm = "ReadAndExecute, Write"}        
                268435456      {$Perm = "FullControl (Subs Only)"}
                2032127        {$Perm = "Full Control"}
                1245631        {$Perm = "Change"}
                default        {$Perm = "None/Other"}
            }

            Write-Output ("Share: {0}, Domain: {1}, User: {2}, Permission: {3}" -f $ShareName, $Domain, $User, $Perm)
            Write-Output ("Share: {0}, Domain: {1}, User: {2}, Permission: {3}" -f $ShareName, $Domain, $User, $Perm) | out-file -FilePath $SMBPathfile -append -encoding ascii            
            
        }
    }
}

# Return To Base
set-location "$BaseFolder"


