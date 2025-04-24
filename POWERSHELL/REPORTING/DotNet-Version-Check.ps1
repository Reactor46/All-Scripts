<#
.SYNOPSIS
Check installed dotnet core runtime version
.DESCRIPTION
.NET Runtime Version Checker (by AlphaBs)
Check installed dotnet runtime version is greater or equal then specific version
.EXAMPLE
DotNet-Version-Check.ps1 -type [type] -minVersion [minVersion]
.EXAMPLE
DotNet-Version-Check.ps1 -type WindowsDesktop -minVersion 6.0.2
.PARAMETER type
Runtime type. can be 'AspNetCore', 'NETCore', 'WindowsDesktop'
AspNetCore: for ASP.NET core apps
NETCore: for console apps
WindowsDesktop: for WPF/WinForm apps
.PARAMETER minVersion
The 
Input format should be 'major.minor.patch' (example: 5.0.2, 6.0.3)
.OUTPUTS
true
The installed dotnet runtime version is greater or equal then `minVersion` (exit code 0)
false
The installed dotnet runtime version is less then `minVersion` (exit code 1)
.NOTES
MIT license
author: AlphaBs(ksi123456ab), https://github.com/AlphaBs
#>


param(
    [string]$type,
    [version]$minVersion,
    [version]$exactVersion
)

function success() {
    Write-Output "true"
    exit 0
}

function fail() {
    Write-Output "false"
    exit 1
}

# there are three runtimes for windows
# https://docs.microsoft.com/en-us/dotnet/core/install/windows?tabs=net60#runtime-information
if ($type -ne "AspNetCore" -and
    $type -ne "NETCore" -and
    $type -ne "WindowsDesktop") {

    Write-Output "Unknown type! Use AspNetCore, NETCore, WindowsDesktop"
    exit 2
}

$stdOutTempFile = "$env:TEMP\$((New-Guid).Guid)"

try {
    Start-Process -FilePath 'dotnet' -ArgumentList '--list-runtimes' -RedirectStandardOutput $stdOutTempFile -Wait -NoNewWindow
    $fileContents = Get-Content $stdOutTempFile
} catch {
    Write-Output "false"
    exit 1
} finally {
    Remove-Item -Path $stdOutTempFile -Force -ErrorAction Ignore
}

if ($minVersion -eq $null -and $exactVersion -eq $null) {
    success
}

$regex = "Microsoft\.${type}\.[a-zA-Z]* ([0-9\.]*)"
$found = $false
foreach($line in $fileContents) {
    if($line -match $regex){
        $version = [version]$matches[1]
        if ($version -ge $minVersion){
            $found = $true
            break
        }
    }
}

if ($found) {
    success
}
else {
    fail
}