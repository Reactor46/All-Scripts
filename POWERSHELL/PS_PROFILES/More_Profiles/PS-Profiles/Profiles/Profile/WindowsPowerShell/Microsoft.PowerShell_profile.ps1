#########################
# Set Path, Update Help #
#########################
Set-Location -Path C:\LazyWinAdmin
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
Get-ChildItem "$env:USERPROFILE\Documents\WindowsPowerShell\Functions\*.ps1" | %{.$_} -ErrorAction SilentlyContinue
##################
# Import-Modules #
##################
Import-Module -Name PowerShellGet
#Import-Module -Name ActiveDirectory
Import-Module -Name PSNmap
Import-Module -Name 7Zip4Powershell
#Import-Module -Name CertificatePS
Import-Module -Name WinSCP
Import-Module UpdateInstalledModule
Start-Job -Name "UpdateHelp" -ScriptBlock { Update-Help -Force } | Out-null
Write-Host "Updating Help in background (Get-Help to check)" -ForegroundColor 'DarkGray'
Start-Job -Name "Update Modules" -ScriptBlock {"Update-InstalledModule"} | Out-Null
Write-host "Updating Modules" -ForegroundColor 'DarkGray'
Write-Host "Custom Scripts Loaded" -ForegroundColor Green
Write-Host "READY!!!!"

$SaveDate = (Get-Date).tostring("MM-dd-yyyy")
#Custom menu that lists currently available functions within the shell repository
function PrintMenu {

	Write-Host(" ----------------------- ")
	Write-Host("'Da List")
	Write-Host(" ----------------------- ")
	Write-Host('Type "GUI" to launch GUI interface!')
	Write-Host("")
	Write-Host("Command             Function")
	Write-Host("-------             --------")
	Write-Host("CheckProcess        Retrieve System Process Information")
	Write-Host("CrossCertRm         Remove Inoperable Certificates")
	Write-Host("Enable              Enable User Account in AD")
	Write-Host("GetSAM              Search For SAM Account Name By Name")
	Write-Host("GPR                 Group Policy (Remote)")
	Write-Host("InstallApplication  Silent Install EXE, MSI, or MSP files")
	Write-Host("JavaCache           Clear Java Cache")
	Write-Host("LastBoot            Get Last Reboot Time")
	Write-Host("NetMSG              On-screen Message For Specified Workstation(s)")
	Write-Host("RDP                 Remote Desktop")
    Write-Host("Reboot              Force Restart")
	Write-Host("RmUserProf          Clear User Profiles")
	Write-Host("SWcheck             Check Installed Software")
	Write-Host("SYS                 All Remote System Info")
	Write-Host("UpdateProfile       Update PowerShell Profile (Will Overwrite Current Version & Any Changes)")
	Write-Host("")
	Write-Host("")
}#End PrintMenu
##########
#PS Ready#
##########
Write-Host ("Get-IPAddress, for, well, ya' know, Yer IP and Stuff") -ForegroundColor Green
Write-Host ("Touch, like Uncle Touchy, but for files") -ForegroundColor Green
Write-Host ("NPP, launch Notepad++") -ForegroundColor Green
Write-Host ("Find, for finding files!") -ForegroundColor Green
Write-Host ("SUDO, for elevated PS. Could use some work, feeling lucky?") -ForegroundColor Green
Write-Host ("Remove-UserProfile, Using delprof2 to delete profiles far and near!") -ForegroundColor Green
Write-Host ("Remove-RemotePrintDrivers, to delete printer drivers far and near!") -ForegroundColor Green
Write-Host ("RDP, for remoting to desktops far and near!.... notice a trend??") -ForegroundColor Green
Write-Host ("Get-LastBoot, find the last time a system rebooted. No exclamation point... its just NOT that exciting..") -ForegroundColor Green
Write-Host ("Get-LoggedOnUser, find someone... loggedon... to a remote system... yeah I know, SysInternals has a tool, but... POWERSHELL!!") -ForegroundColor Green
Write-Host ("Get-HotFixes, Yes.. it finds hotfixes installed on remote systems.") -ForegroundColor Green
Write-Host ("CHKProcess, check running processes remotely.") -ForegroundColor Green
Write-Host ("WHOIS, do a whois lookup!!! WhoIs -domain power-shell.com") -ForegroundColor Green
Write-Host ("Custom Scripts Loaded") -ForegroundColor Green
Write-Host ("READY!!!!")

#Set Home Path
(Get-PSProvider 'FileSystem').Home = "C:\LazyWinAdmin"
Get-MOTD
