
#Module Browser Begin
#Version: 1.0.0
Add-Type -Path 'C:\Program Files (x86)\Microsoft Module Browser\ModuleBrowser.dll'
$moduleBrowser = $psISE.CurrentPowerShellTab.VerticalAddOnTools.Add('Module Browser', [ModuleBrowser.Views.MainView], $true)
$psISE.CurrentPowerShellTab.VisibleVerticalAddOnTools.SelectedAddOnTool = $moduleBrowser
#Module Browser End
Import-Module C:\Scripts\Functions\EzAdmin.psm1 -Force 
Import-Module -Name PSReadLine
$MaximumHistoryCount = 10000
$PSDefaultParameterValues["Out-File:Encoding"]="utf8"
Set-PSReadlineKeyHandler -Chord Tab -Function MenuComplete
cd\
Clear-Host
