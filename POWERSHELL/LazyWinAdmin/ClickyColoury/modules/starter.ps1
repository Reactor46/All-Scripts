# api: ps
# type: main
# title: ClickyColoury + TextyTypey - GUI+CLI tool frontend
# description: Convenience invocation of various Powershell and CMD scripts
# url: http://fossil.include-once.org/clickycoloury/
# version: 0.8.0
# depends: menu, wpf, clipboard
# category: misc
# config:
#   { name: cfg.gridview, type: select, value: Format-Table, select: Format-Table|Out-GridView, description: default table display mode }
#   { name: cfg.cli, type: bool, value: 0, description: Start console (CLI) version per default? }
#   { name: cfg.cached, type: bool, value: 0, description: Use CLIXML script cache on startup }
#   { name: debug, type: bool, value: 0, description: Powershell-internal } 
# author: mario
# license: MITL
# priority: core
# status: testing
#
# Note that this is the WindowsPresentationForm version, but also renders a
# classic -CLI menu otherwise. Utilizes:
#
#   · wpf.psm1 = graphical toolkit features
#   · menu.psm1 = mostly CLI features
#   · clipboard.psm1 = for the HTML output
#
# Whereas scripts (and plugins) reside in tools*/*.ps1
#
# Starting up with `-cli` parameter should yield the text version.
#
# Only works with powershell.exe -Version 2.0 -STA -File ... at the moment.
#


#-- params
[CmdletBinding()]
Param(
    [switch]$CLI = $false
)

#-- config
$global:cfg = @{
    domain = "ContosoCORP" #(Get-ADDomain).Netbiosname
    threading = 1
    autoclear = 300
    exchange = @{
        ConfigurationName = "Microsoft.Exchange"
        ConnectionUri = "http://LASEXCH02/PowerShell/"
    }
    main = @{} 
    curr_script_fn = $MyInvocation.MyCommand.Path
    multitool_base = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
    user_config_fn  = "$env:APPDATA\multitool\config.ps1"
    user_plugins_d  = "$env:APPDATA\multitool"
    tools_cache_fn = "./data/tools.cache.clixml"
    tool_dirs = @("./tools/*/*.ps1")  # add custom dirs here!
}
$global:plugins = @{
    "init" = @()
    "before" = @()
    "after" = @()
    "menu" = @()
}
#-- load user config
cd ($cfg.multitool_base)
if (Test-Path ($cfg.user_config_fn)) {
    . ($cfg.user_config_fn)
}
#$cfg|FL

#-- restart if not in single-thread-apartment mode
if (-not [System.Management.Automation.Runspaces.Runspace]::DefaultRunSpace.ApartmentState -eq "STA") {
#    powershell.exe -Version 2.0 -STA -ExecutionPolicy unrestricted -File $curr_script_fn
#    break 2
}

#-- modules
$global:GUI = [hashtable]::Synchronized(@{Host=$Host})
Import-Module -DisableNameChecking "$($cfg.multitool_base)\modules\wpf.psm1"
Import-Module -DisableNameChecking "$($cfg.multitool_base)\modules\menu.psm1"
Import-Module -DisableNameChecking "$($cfg.multitool_base)\modules\clipboard.psm1"
if (!(Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory }

#-- post init
$cfg.main = (Extract-PluginMeta $cfg.curr_script_fn)
#$global:GUI|FL



#-- predefined menu entries
$menu = @(
)
#-- load cache?
if ($cfg.cached -and (Test-Path ($cfg.tools_cache_fn))) {
    $menu = Import-CliXml ($cfg.tools_cache_fn)
}
#-- add menu entries from scripts
else {
    $menu += @(Get-Item ($cfg.tool_dirs) | % { Extract-PluginMeta $_ } | ? { $_.api -and $_.type -and $_.title })
}
#-- add user plugins
if (Test-Path ($cfg.user_plugins_d)) {
    $menu += @(Get-Item "$($cfg.user_plugins_d)/*.ps1" | % { Extract-PluginMeta $_ } | ? { $_.type -and $_.api -eq "multitool" })
}
#-- run `type:init` plugins here
$menu | ? { $_.type -eq "init" -and $_.fn } | % { . $_.fn }


#-- CLI mode
if ($CLI) {
    Init-Screen
    $menu = $menu + @{key="m|menu"; category="extras"; title="print menu"; func='Print-Menu $menu'}
    Print-Menu $menu
    Process-Menu -Menu $menu -Prompt "TextyTypey"
}
#-- WPF multi-tool
else {
    echo "Starting GUI version..."
    $shell = Start-Win (Sort-Menu $menu)
}

#-- cleanup
if ($Debug) {
    Remove-Module wpf
    $Error
}
