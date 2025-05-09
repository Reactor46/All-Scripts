# Report on Plugins.ps1

# List plugins and whether they're enabled or disabled by Select-Plugins.ps1

# Start of Settings
# List Enabled Plugins First
$ListEnabledPluginsFirst=$True
# End of Settings

# Changelog
## 1.0 : Initial Release
## 1.1 : Added Changelog
## 1.2 : Made plugin code consistent with Select-Plugins.ps1

$Title = "Report on Plugins"
$Author = "Phil Randal"
$PluginVersion = 1.1
$Header =  "Plugins Report"
$Comments = "Plugins in alphabetical order"
$Display = "Table"
$PluginCategory = "vCheck"

$PluginPath = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)
If ($PluginPath -notmatch 'plugins$') {
  $PluginPath += "\Plugins"
}
$plugins=get-childitem -Path $PluginPath | where {$_.name -match '.*\.ps1(?:\.disabled|)$'} |
   Select Name, 
          @{Label="Plugin";expression={$_.Name -replace '(.*)\.ps1(?:\.disabled|)$', '$1'}},
          @{Label="Enabled";expression={$_.Name -notmatch '.*\.disabled$'}}

If ($ListEnabledPluginsFirst) {
  $Plugins |
    Sort -property @{Expression="Enabled";Descending=$true}, @{Expression="Plugin";Descending=$false} |
    Select Plugin, Enabled
  $Comments = "Plugins in alphabetical order, enabled plugins listed first"
} Else {
  $Plugins | Sort Plugin | Select Plugin, Enabled
}
$Plugins = $null
$Comments = "Plugins in alphabetical order"
