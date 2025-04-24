[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '')]
param()

$stopwatch = [system.diagnostics.stopwatch]::StartNew()
$enableTiming = $false
# pwsh -noprofile
# . "C:\Users\alexko\OneDrive - Microsoft\Documents\PowerShell\profile.ps1"

function tm($info = "=>")
{
    if( $enableTiming )
    {
        Write-Host "$($stopwatch.ElapsedMilliseconds / 1000) $info"
        $stopwatch.Restart()
    }
}

# Powershell behavior setup
$global:Profile = $PSCommandPath
$global:MaximumHistoryCount = 1024
$env:PSModulePath += ";$PSScriptRoot\Modules"
tm init

# Default command arguments
$PSDefaultParameterValues["Get-Command:All"] = $true
tm defaults

# Was fixed in Windows 10
if( [Environment]::OSVersion.Version.Major -lt 10 )
{
    Update-FormatData -PrependPath "$PSScriptRoot\Format.Custom.ps1xml"
}

Set-Alias new New-Object
Set-Alias m Measure-Object
tm alias

# Environment setup
$addToPath =
"c:\tools\LinqPad5\",
"c:\tools\xts\",
"C:\Program Files\Beyond Compare 4\",
"C:\Program Files (x86)\WinDirStat\"

$env:Path += ";" + ($addToPath -join ";")
tm environment

# Additional setup
. $PSScriptRoot\Scripts\Playground.ps1
tm playground

. $PSScriptRoot\Scripts\Load-Functions.ps1
Remove-Variable proc -ea Ignore # hides pro<tab> = profile
tm func

. $PSScriptRoot\Scripts\Initialize-Computer.ps1
tm comp

. $PSScriptRoot\Scripts\Initialize-PsReadLine.ps1
tm psreadline

. $PSScriptRoot\Scripts\Initialize-Prompt.ps1
tm prompt


# That's hacky... but it can dot script other script here
if( -not (Test-Path "$oneDriveMicrosoft\Projects\ProtectedPlayground.ps1") )
{
    return
}
. "$oneDriveMicrosoft\Projects\ProtectedPlayground.ps1"
tm ProtectedPlayground

$enableTiming = $true