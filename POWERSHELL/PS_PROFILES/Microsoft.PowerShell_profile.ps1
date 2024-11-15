#########################
# Set Path, Update Help #
#########################
Set-Location -Path C:\LazyWinAdmin
#Start-Job -Name "UpdateHelp" -ScriptBlock { Update-Help -Force } | Out-null
#Write-Host "Updating Help in background (Get-Help to check)" -ForegroundColor 'DarkGray'
#Start-Job -Name "Update Modules" -ScriptBlock {"$PSScriptRoot\Scripts\Check-ModuleUpdates.ps1"} | Out-Null
#Write-host "Updating Modules" -ForegroundColor 'DarkGray'
Get-PSSession | Remove-PSSession
New-Alias touch Set-FileTime
################
# Shell Colors #
################
#$shell = $Host.UI.RawUI
#$shell.BackgroundColor="Black"
#$shell.ForegroundColor="Green"
################################
# History and Various Settings #
################################
#Import-Module -Name PSReadLine
#$MaximumHistoryCount = 10
#$PSDefaultParameterValues["Out-File:Encoding"]="utf8"
#Set-PSReadlineKeyHandler -Chord Tab -Function MenuComplete
#Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
#########################
# Check Admin Elevation #
#########################
$WindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$WindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($WindowsIdentity)
$Administrator = [System.Security.Principal.WindowsBuiltInRole]::Administrator
$IsAdmin = $WindowsPrincipal.IsInRole($Administrator)

################################
# AutoLoad Additional Functions#
################################
Get-ChildItem "C:\Users\john.advisor\Documents\WindowsPowerShell\Functions\*.ps1" | %{.$_} -ErrorAction SilentlyContinue
##################
# Import-Modules #
##################
Import-Module -Name ActiveDirectory
Import-Module -Name ActiveDirectoryTools
Import-Module -Name Add-on.ModuleManagement
Import-Module -Name PSNmap
Import-Module -Name dbatools

#Write-Host "Custom Scripts Loaded" -ForegroundColor Green
#Write-Host "READY!!!!"
