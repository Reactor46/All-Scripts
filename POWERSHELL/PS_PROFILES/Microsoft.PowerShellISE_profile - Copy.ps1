#########################
# Set Path, Update Help #
#########################
Set-Location -Path C:\LazyWinAdmin
#Start-Job -Name "UpdateHelp" -ScriptBlock { Update-Help -Force } | Out-null
#Write-Host "Updating Help in background (Get-Help to check)" -ForegroundColor 'DarkGray'
#Start-Job -Name "Update Modules" -ScriptBlock {"$PSScriptRoot\Scripts\Check-ModuleUpdates.ps1"} | Out-Null
#Write-host "Updating Modules" -ForegroundColor 'DarkGray'
Get-PSSession | Remove-PSSession
################################
# History and Various Settings #
################################
Import-Module -Name PSReadLine
$MaximumHistoryCount = 10000
$PSDefaultParameterValues["Out-File:Encoding"]="utf8"
Set-PSReadlineKeyHandler -Chord Tab -Function MenuComplete
#########################
# Application Variables #
#########################
#Add-PathVariable "C:\OpenSSL-Win32\bin"
#Add-PathVariable "${env:ProgramFiles}\OpenSSH-Win64"
#Add-PathVariable "${env:ProgramFiles}\7-Zip"
###################
# Trust PSGallery #
###################
#Get-PackageProvider -Name NuGet -ForceBootstrap
#Set-PackageSource -Name NuGet -Trusted
#Get-PackageProvider -Name Chocolatey -ForceBootstrap
#Set-PackageSource -Name Chocolatey -Trusted
#Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
#########################
# Check Admin Elevation #
#########################
$WindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$WindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($WindowsIdentity)
$Administrator = [System.Security.Principal.WindowsBuiltInRole]::Administrator
$IsAdmin = $WindowsPrincipal.IsInRole($Administrator)
################################
# Add Verbose and Debug to ISE #
################################
$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Clear()
$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add('Run with -Verbose', { Invoke-Expression -Command ". '$($psISE.CurrentFile.FullPath)' -Verbose" }, 'Ctrl+F5') | Out-Null
$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add('Run with -Debug',   { Invoke-Expression -Command ". '$($psISE.CurrentFile.FullPath)' -Debug" }, 'Ctrl+F6') | Out-Null
####################
# Static Functions #
####################
# Touch
function touch { $args | foreach-object {write-host > $_} }
# Notepad++
function NPP { Start-Process -FilePath "${Env:ProgramFiles(x86)}\Notepad++\Notepad++.exe" }#-ArgumentList $args }
# Find File
function findfile($name) {
	ls -recurse -filter "*${name}*" -ErrorAction SilentlyContinue | foreach {
		$place_path = $_.directory
		echo "${place_path}\${_}"
	}
}
# RM -RF
function rm-rf($item) { Remove-Item $item -Recurse -Force }
# SUDO
function sudo(){
	Invoke-Elevated @args
}
# SED
function sed($file, $find, $replace){
	(Get-Content $file).replace("$find", $replace) | Set-Content $file
}
# SED-Recursive
function sed-recursive($filePattern, $find, $replace) {
	$files = ls . "$filePattern" -rec # -Exclude
	foreach ($file in $files) {
		(Get-Content $file.PSPath) |
		Foreach-Object { $_ -replace "$find", "$replace" } |
		Set-Content $file.PSPath
	}
}
# Grep
function grep($regex, $dir) {
	if ( $dir ) {
		ls $dir | select-string $regex
		return
	}
	$input | select-string $regex
}

# Grep W/Values
function grepv($regex) {
	$input | ? { !$_.Contains($regex) }
}
# Which
function which($name) {
	Get-Command $name | Select-Object -ExpandProperty Definition
}
# Cut
function cut(){
	foreach ($part in $input) {
		$line = $part.ToString();
		$MaxLength = [System.Math]::Min(200, $line.Length)
		$line.subString(0, $MaxLength)
	}
}
# Search Text Files
function Search-AllTextFiles {
    param(
        [parameter(Mandatory=$true,position=0)]$Pattern, 
        [switch]$CaseSensitive,
        [switch]$SimpleMatch
    );

    Get-ChildItem . * -Recurse -Exclude ('*.dll','*.pdf','*.pdb','*.zip','*.exe','*.jpg','*.gif','*.png','*.ico','*.svg','*.bmp','*.psd','*.cache','*.doc','*.docx','*.xls','*.xlsx','*.dat','*.mdf','*.nupkg','*.snk','*.ttf','*.eot','*.woff','*.tdf','*.gen','*.cfs','*.map','*.min.js','*.data') | Select-String -Pattern:$pattern -SimpleMatch:$SimpleMatch -CaseSensitive:$CaseSensitive
}
# Add to Zip
function AddTo-7zip($zipFileName) {
    BEGIN {
        #$7zip = "$($env:ProgramFiles)\7-zip\7z.exe"
        $7zip = Find-Program "\7-zip\7z.exe"
		if(!([System.IO.File]::Exists($7zip))){
			throw "7zip not found";
		}
    }
    PROCESS {
        & $7zip a -tzip $zipFileName $_
    }
    END {
    }
}
# Connect to Exchange
Function GoGo-PSExch {

    param(
        [Parameter( Mandatory=$false)]
        [string]$URL="LASEXCH01"
    )
    
    #$Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

    $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos #-Credential #$Credentials

    Import-PSSession $ExOPSession -AllowClobber

}
################################
# AutoLoad Additional Functions#
################################
Get-ChildItem "C:\Users\jbattista.advisor\Documents\WindowsPowerShell\Functions\*.ps1" | %{.$_} -ErrorAction SilentlyContinue
##################
# Import-Modules #
##################
#Import-Module -Name ISEThemes
#Import-Module -Name ISEColorTheme.Cmdlets
Import-Module -Name ISERemoteTab
Import-Module -Name ISEScriptingGeek
Import-Module -Name ISEScriptAnalyzerAddOn
Import-Module -Name PowerShellGet
Import-Module -Name ActiveDirectory
Import-Module -Name Add-on.ModuleManagement
Import-Module -Name PSNmap
Import-Module -Name ProtectedData
Import-Module -Name FormsGUI
Import-Module -Name AssetInventory
Import-Module -Name Chalk
Import-Module -Name ChocolateyGet
Import-Module -Name CimInventory
Import-Module -Name CimSession
Import-Module -Name CimSweep
Import-Module -Name CliMenu
Import-Module -Name ConfigExport
Import-Module -Name CompareComputer
Import-Module -Name CustomizeWindows10
Import-Module -Name EnhancedHTML2
Import-Module -Name GPODoc
Import-Module -Name ReportHTML

##############
# Local Repo #
##############
$Path = '\\laspshost\Scripts\Repository\jbattista'
$repo = @{
    Name = 'C1B-PSHOST'
    SourceLocation = $Path
    PublishLocation = $Path
    InstallationPolicy = 'Trusted'
}
Register-PSRepository @repo -ErrorAction SilentlyContinue
##########
#PS Ready#
##########
Write-Host "Get-IPAddress, for, well, ya' know, Yer IP and Stuff" -ForegroundColor Green
Write-Host "Touch, like Uncle Touchy, but for files" -ForegroundColor Green
Write-Host "NPP, launch Notepad++" -ForegroundColor Green
Write-Host "Find, for finding files!" -ForegroundColor Green
Write-Host "SUDO, for elevated PS. Could use some work, feeling lucky?" -ForegroundColor Green
Write-Host "Remove-UserProfile, Using delprof2 to delete profiles far and near!" -ForegroundColor Green
Write-Host "Remove-RemotePrintDrivers, to delete printer drivers far and near!" -ForegroundColor Green
Write-Host "RDP, for remoting to desktops far and near!.... notice a trend??" -ForegroundColor Green
Write-Host "Get-LastBoot, find the last time a system rebooted. No exclamation point... its just NOT that exciting.." -ForegroundColor Green
Write-Host "Get-LoggedOnUser, find someone... loggedon... to a remote system... yeah I know, SysInternals has a tool, but... POWERSHELL!!" -ForegroundColor Green
Write-Host "Get-HotFixes, Yes.. it finds hotfixes installed on remote systems." -ForegroundColor Green
Write-Host "CHKProcess, check running processes remotely." -ForegroundColor Green
Write-Host "WHOIS, do a whois lookup!!! WhoIs -domain power-shell.com" -ForegroundColor Green
Write-Host "Custom Scripts Loaded" -ForegroundColor Green
Write-Host "READY!!!!"

Get-MOTD