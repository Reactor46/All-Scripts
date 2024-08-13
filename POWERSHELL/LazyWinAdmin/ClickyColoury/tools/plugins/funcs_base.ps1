# encoding: utf-8
# api: ps
# title: functions
# description: Some utility functions for scripts
# version: 1.0
# type: init
# category: misc
# hidden: 1
# priority: optional
#
# Defines:
#  · PSExec                  → wrap psexec.exe
#  · Check-PSExecResult      → eval $LASTEXITCODE
#  · Invoke-ExchangeCommand  → simpler remoting
#  · Open-RemoteRegistry     → -hive LocalMachine -machine A0004915 -path SOFTWARE
#


#-- check $LASTEXITCODE from psexec invocations
function Check-PSExecResult {
    Param($err=$LASTEXITCODE)
    switch ($err) {
        0 { Write-Host -f Yellow "✔ okay" }
        1 { Write-Host -f Yellow "✔ okay" }
        2 { Write-Host -f Red "✘ failed" }
        default { Write-Host -f Gray "➗ Exitcode=$err" }
    }
}
#-- might as well define a PSExec wrapper then...
function PSExec {
    [CmdletBinding()]
    Param($machine, [Parameter(ValueFromRemainingArguments=$true)]$cmd)
    $cmd = (($cmd | % { if ($_ -match ' ') { '"'+$_+'"' } else { $_ }}) -join ' ')

    #Write-Host -f DarkYellow "❏ PSExec($machine $cmd)"
    Invoke-Expression "& PSExec.exe $machine $cmd 2>&1" | Select -Skip 3 | Write-Host -b DarkGray
                                                   # 
    [void](Check-PSExecResult $LASTEXITCODE)
}


#-- default session (just started once, then kept in memory)
function Import-ExchangeSession {
    $params = $cfg.exchange
    if (! (Test-Path function:Get-Mailbox)) {
        $null = ( Write-Host -f Green "❏ Exchange connection..." )
        $global:Exchange_Session = New-PSSession @params
        $null = Import-PSSession -Session $global:Exchange_Session
    }
    else {
        $null = ( Write-Host -f DarkGray "✔ Exchange conn active." )
    }
}

#-- via WMI _ComputerSystem or looking up Owner of Explorer _Process
function Get-CurrentUser {
    Param($machine)
    if (($W = GWMI win32_computersystem -comp $machine) -and ($W.username)) {
        $r = $W.username
    }
    elseif ($W = GWMI Win32_Process -ComputerName $machine -filter "name='Explorer.exe'") {
        $r = ($w | % { $_.getOwner().user } | ? { $_ -notmatch "^SYSTEM$" })
        if ($r -is [array]) { $r = $r[0] }
    }
    else {
        $r = "nobody"
    }
    return $r -replace "^\w+\\(?=\w+)",""
}

#-- via WMI _userAccount
function Get-UserSID {
    Param($user)
    ([WMI]"win32_userAccount.Domain='$($cfg.domain)',Name='$user'").SID
}


#-- remote registry
function Open-RemoteRegistry {
    <#
    .SYNOPSIS
        Open remote registry tree
    .DESCRIPTION
        Establish a Win32.Registry connection to remote machine. Does not retrieve Leafes/Values itself.
    .PARAMETER  Path
        Can either be a full path such as "\\HOSTNAME\HKLM\SOFTWARE\Windows"
        Or just the regpath "SOFTWARE\Windows" when both -Hive and -Machine are given
    .PARAMETER  Hive
        If no full -Path given, should name "HKLM", "HKCR", or "HKCU" (current user is looked up automatically).
    .PARAMETER  Machine
        If no full -Path given, lists the remote hostname to connect to.
    .EXAMPLE
        $R = Open-RemoteRegisty "\\localhost\HKLM\SW\WindowsCurrentControlSet"
        $R = Open-RemoteRegisty -Machine "localhost" -Hive "HKLM" -Path "SW\WindowsCurrentControlSet"
    #>
    Param(
        $path = $null,            # preferred: full specifier "\\HOSTNAME\HKLM\RegPath" (host+hive+path; no leaf/value)
        $hive = "LocalMachine",
        $machine = $null,
        $writemode = $true,
        [switch]$silent = $false
    )
 
    #-- combine path if it starts with "\\" two backslashes
    if ((!$machine) -and ($path -match "^\\\\(\w+)\\(\w+)\\(.+)$")) {
        $machine = $matches[1]
        $hive = $matches[2]
        $path = $matches[3]
    }

    #-- hive aliases
    $hive = switch -regex ($hive) {
        ".*LM" { "LocalMachine" }
        ".*CR" { "ClassesRoot" }
        ".*CU" {
             $path = (Get-UserSID (Get-CurrentUser $machine)) + "\" + $path
             [Microsoft.Win32.RegistryHive]::Users
        }
        ".*USERS" { [Microsoft.Win32.RegistryHive]::Users }
    }
    if (!$hive) {
        $hive = "LocalMachine"
    }

    #-- open
    if (!$silent) {
        $null = ( Write-Host -f DarkGray  "❏ Remote registry connection [$machine\$hive]..." )
    }
    $R = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("$hive", $machine)
    if (!$silent) {
        $null = ( Write-Host -f DarkGray  "❏ Open subkey [$path]..." )
    }
    return $R.OpenSubKey($path, $writemode)
}

#-- shortcuts
function Set-RemoteRegistry {
    Param($path = "\\localhost\HKLM\SOFTWARE\Etc\Key", $value="", $type="String")
    if ($path -match "^\\\\(\w+)\\(\w+)\\(.+)\\([^\\]+)$") {
        $R = Open-RemoteRegistry -machine $matches[1] -hive $matches[2] -path $matches[3] -Silent
        $R.setValue($matches[4], $value, $type)
    }
}
function Get-RemoteRegistry {
    Param($path = "\\localhost\HKLM\SOFTWARE\Etc\Key")
    if ($path -match "^\\\\(\w+)\\(\w+)\\(.+)\\([^\\]+)$") {
        $R = Open-RemoteRegistry -machine $matches[1] -hive $matches[2] -path $matches[3] -writemode $false -Silent
        return $R.getValue($matches[4])
    }
}
