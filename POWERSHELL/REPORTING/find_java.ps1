﻿<#
.SYNOPSIS
Locates Java versions and optionally sets JAVA_HOME and JRE_HOME.
.DESCRIPTION
The Find-Java function uses PATH, JAVA_HOME, JRE_HOME and Windows Registry to retrieve installed Java versions.
If run with no options, the first Java found is printed on the console.
.PARAMETER Vendor
Selects Java vendor, currently supports Oracle, OpenJDK and IBM. Defaults to Any.
.PARAMETER Architecture
What processor architecture to match. Valid options are 32, 64, Match and All. Match detects what integer size is used for the PowerShell process, and matches the architecture. If Wow64 is available, 64 bit versions of Java are selected first.
.PARAMETER Package
Selects Java package, currently supports JRE, JDK and Server-JRE. Defaults to Any.
.PARAMETER Version
Selects Java specification version, from 1.6 to 1.8, 9 to 11, or latest found. Defaults to Latest.
.PARAMETER SetSession
Sets JAVA_HOME and JRE_HOME for current session (process).
.PARAMETER SetUser
Permanently sets JAVA_HOME and JRE_HOME for current user.
.PARAMETER SetSystem
Permanently sets JAVA_HOME and JRE_HOME for the system. Requires administrator rights.
.PARAMETER Output
Outputs found Java properties. Valid options are AsEnv, None, All. Defaults to None.
.EXAMPLE
Sets JAVA_HOME and JRE_HOME for the latest version of Oracle JRE
Find-Java -Vendor Oracle -Package JRE -SetSession
.EXAMPLE 
Prints all Javas and their properties as JSON
Find-Java -Output All | ConvertTo-JSON
#>
Param(
    [Parameter(HelpMessage = "Preferred Java vendor")][String][ValidateSet("Oracle", "OpenJDK", "IBM", "Zulu", "Any")]$Vendor = "Any",
    [Parameter(HelpMessage = "Preferred Java architecture")][String][ValidateSet("32", "32bit", "64", "64bit", "Match", "All")]$Architecture = "Match",
    [Parameter(HelpMessage = "Preferred Java package")][String][ValidateSet("JRE", "JDK", "Server-JRE", "Any")]$Package = "Any",
    [Parameter(HelpMessage = "Preferred Java version")][String][ValidateSet("1.6", "1.7", "1.8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "Latest")]$Version = "Latest",
    [Switch]$SetSession,
    [Switch]$SetUser,
    [Switch]$SetSystem,
    [Parameter(HelpMessage = "Outputs found Java")][String][ValidateSet("AsEnv", "None", "All", "JSON")]$Output = "None"
)

function is_admin {
    $admin = [Security.Principal.WindowsBuiltInRole]::Administrator
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    ([Security.Principal.WindowsPrincipal]($id)).IsInRole($admin)
}

function is_win32() {
    return [IntPtr]::size -eq 4
}

function is_win64() {
    return [IntPtr]::size -eq 8
}

function is_wow() {
    return (is_win32) -and (test-path env:\PROCESSOR_ARCHITEW6432)
}

function get_package_name([String]$javahome) {
    if ((Test-Path "$(Split-Path $javahome)\jre") -or (Test-Path "$javahome\jre")) {
        $realhome = Split-Path $javahome
        if (((Test-Path "$realhome\src.zip") -or (Test-Path "$realhome\db")) -and (Test-Path "$realhome\bin\java.exe")) {
            return "JDK"
        }
        elseif ((Test-Path "$realhome\lib\tools.jar") -and (Test-Path "$realhome\bin\java.exe")) {
            return "Server-JRE"
        }
    }
    return "JRE"
}

function java_properties([String]$java) {
    $file = New-TemporaryFile
    Write-Verbose "Getting properties by executing ""$java"""
    If (-not (Test-Path $java)) {
        Write-Verbose "Could not find ""$java"""
        return
    }
    [Void](Start-Process $java -NoNewWindow -Wait -ErrorAction Stop -ArgumentList "-XshowSettings:properties -version" -RedirectStandardError $file.FullName)
    [System.Collections.SortedList]$properties = [Hashtable]@{}
    foreach ($line in Get-Content $file.FullName) {
        Write-Debug $line
        if ($line -match '[\s]{4}(?<property>[\w.]+)\s=\s(?<value>.*)') {
            $properties.Add($Matches["property"], $Matches["value"])
            $last = $Matches["property"]
        }
        elseif (($line -match "[\s]{8}(?<value>.*)") -or ((($last -eq "java.vm.info") -or ($last -eq "java.fullversion")) -and ($line -match "(?<value>.*)"))) {
            $value = $properties.Get_Item($last)
            if ($value -is [String]) {
                $properties.Set_Item($last, [Array]@( $value ))
            }
            if ($value -is [Array]) {
                $properties.$last += $Matches["value"]
            }
        }
        else {
            Write-Debug "Not found: $line"
        }
    }
    [Void](Remove-Item $file)
    $properties.Add("custom.package.name", $(get_package_name $properties."java.home" ))
    $jrehome = $javahome = $properties."java.home"
    if ($properties."custom.package.name" -ne "JRE") {
        $javahome = Split-Path $jrehome
        if (-not (Test-Path "$javahome\bin\java.exe")) {
            $javahome = $jrehome
        }
    }
    $properties.Add("custom.java.home", $javahome)
    $properties.Add("custom.jre.home", $jrehome)
    $properties.Add("custom.java.path", $java)
    $properties.Add("custom.vendor", "Unknown")
    if (($properties."java.vendor" -eq "Oracle Corporation") -and ($properties."java.vm.name" -notlike "OpenJDK*")) {
        $properties.Set_Item("custom.vendor", "Oracle")
    }
    if (($properties."java.vendor" -eq "Oracle Corporation") -and ($properties."java.vm.name" -like "OpenJDK*")) {
        $properties.Set_Item("custom.vendor", "OpenJDK")
    }
    if ($properties."java.vendor" -eq "IBM Corporation") {
        $properties.Set_Item("custom.vendor", "IBM")
    }
    if ($properties."java.vendor" -eq "Azul Systems, Inc.") {
        $properties.Set_Item("custom.vendor", "Zulu")
    }
    Write-Verbose "Found Java: $($properties.'custom.vendor') $($properties.'custom.package.name')"
    return $properties
}

function java_from_env([String]$envvar) {
    Write-Verbose "Checking ""$envvar"""
    if (Test-Path env:$envvar) {
        Write-Verbose "Checking for Java in ""$((Get-Item env:$envvar).Value)"""
        return $(java_properties "$((Get-Item env:$envvar).Value)\bin\java.exe")
    }
    Write-Verbose "No Java found in $envvar"
}

function java_from_path {
    Write-Verbose "Searching for ""java.exe"" in PATH"
    if ($java = Get-Command "java.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Definition) {
        Write-Verbose "Checking for Java in ""$(Split-Path $java)"""
        return $(java_properties $java)
    }
    Write-Verbose "No Java found in PATH"
}

function java_from_full_path {
    Write-Verbose "Searching PATH"
    $javas = [Array]@()
    foreach ($path in $env:PATH.split(";")) {
        If (Test-Path "$path\java.exe") {
            Write-Verbose "Checking for Java in ""$path"""
            $javas += $(java_properties "$path\java.exe")
        }
    }
    if ($javas.Length -eq 0) { Write-Verbose "No Java found in PATH" }
    return $javas
}

function java_from_registry([Switch]$wow = $false) {
    $javas = [Array]@()
    $prefix = "HKLM:\\SOFTWARE"
    if (is_win32) { $arch = "32" }
    if (is_win64) { $arch = "64" }
    if ($wow -eq $true) {
        $prefix += "\\WOW6432Node"
        $arch = "32"
    }
    $ErrorActionPreference = "SilentlyContinue"
    foreach ($path in @("$prefix\\JavaSoft\\Java Development Kit", "$prefix\\JavaSoft\\Java Runtime Environment", "$prefix\\IBM\\Java Development Kit", "$prefix\\IBM\\Java Runtime Environment")) {
        Write-Verbose "Searching Windows registry at ""$path"""
        Get-ChildItem -Path $path | ForEach-Object {
            Write-Verbose "Checking for Java in ""$(($_ | Get-ItemProperty -name JavaHome).JavaHome)"""
            $javas += @(java_properties "$(($_ | Get-ItemProperty -name JavaHome).JavaHome)\bin\java.exe")
        }
    }
    if ($javas.Length -eq 0) { Write-Verbose "No Java found in Windows registry" }
    return $javas
}

function sort_javas([Array]$javas) {
    return [Array]@( $javas | Sort-Object -Property "java.version", @{Expression={ if ($_."custom.vendor" -eq "Oracle") { 1 } elseif ($_."custom.vendor" -eq "OpenJDK") { 2 } elseif ($_."custom.vendor" -eq "IBM") { 3 } else { 4 }}; Ascending=$true }, "com.ibm.oti.jcl.build" -Descending )
}

function select_architecture([Array]$javas, [String]$arch) {
    switch ($arch) {
        "64bit" { $arch = "64" }
        "32bit" { $arch = "32" }
        "Match" {
            if (is_win32) { $arch = "32" }
            if ((is_win64) -or (is_wow)) { $arch = "64" }
        }
        "Any" {
            return $javas
        }
    }
    return [Array]@( $javas | Where-Object { $_."sun.arch.data.model" -eq "$arch"} )
}

function select_vendor([Array]$javas, [String]$vendor) {
    switch ($vendor) {
        "Oracle" { return [Array]@( $javas | Where-Object { $_."custom.vendor" -eq "Oracle" } ) }
        "OpenJDK" { return [Array]@( $javas | Where-Object { $_."custom.vendor" -eq "OpenJDK" } ) }
        "IBM" { return [Array]@( $javas | Where-Object { $_."custom.vendor" -eq "IBM" } ) }
    }
    return $javas
}

function select_package([Array]$javas, [String]$package) {
    if ($package -ne "Any") {
        return [Array]@( $javas | Where-Object { $_."custom.package.name" -eq $package} )
    }
    return $javas
}

function select_version([Array]$javas, [String]$version) {
    if ($version -ne "Latest") {
        return [Array]@( $javas | Where-Object { $_."java.specification.version" -eq $version} )
    }
    return $javas
}

function set_environment([String]$target, [String]$javahome, [String]$jrehome) {
    Write-Verbose "Setting JAVA_HOME for ""$target"" to ""$javahome"""
    [Environment]::SetEnvironmentVariable("JAVA_HOME", $javahome, $target)
    Write-Verbose "Setting JRE_HOME for ""$target"" to ""$jrehome"""
    [Environment]::SetEnvironmentVariable("JRE_HOME", $jrehome, $target)
}

$javas = [Array]@()
:find_java foreach ($iex in @( 'java_from_path', 'java_from_env "JRE_HOME"', 'java_from_env "JAVA_HOME"', 'java_from_registry', 'java_from_registry -wow', 'java_from_full_path' )) {
    $javas += [Array]@( Invoke-Expression $iex )
    $javas = @( select_architecture $javas $Architecture )
    $javas = @( select_package $javas $Package )
    $javas = @( select_vendor $javas $Vendor )
    $javas = @( select_version $javas $Version )
    if ($javas.Length -eq 0) { continue }
    if (($Output -eq "AsEnv") -or ($SetSession) -or ($SetUser) -or ($SetSystem)) {
        break find_java
    }
}

if ($javas.Length -eq 0) {
    throw "ERROR: Java not found!"
}

$javas = @( sort_javas $javas )

$javahome = $javas[0]."custom.java.home"
$jrehome = $javas[0]."custom.jre.home"

if ($SetSession) {
    set_environment "Process" $javahome $jrehome
}
if ($SetUser) {
    set_environment "User" $javahome $jrehome
}
if ($SetSystem) {
    if (is_admin) {
        set_environment "Machine" $javahome $jrehome
    }
    else {
        Write-Error "ERROR: You need Administrator rights to set system environment variables!"
    }
}
switch ($Output) {
    "AsEnv" {
        "`$env:JAVA_HOME = ""$javahome"""
        "`$env:JRE_HOME = ""$jrehome"""
    }
    "All" {
        $javas
    }
    "JSON" {
        $javas | ConvertTo-Json
    }
}