# ************************************************************************** #
# ||    	           ~~ AllUsersCurrentHost (Console) ~~                || #
# ||                         Written by Noah Hopkins                      || #
# ||                       Last edited: October, 2018                     || #
# ************************************************************************** #

# NOTE: This profile will affect all Powershell console instances, but not 
# Powershell ISE.

# Purpose: Global modification of console UI/colors


# $Profile.AllUsersAllHosts       (All)     # Global aliases
# $Profile.CurrentUserAllHosts	  (All)     # Does not exist, Noah-specific aliases (none)
# $Profile.AllUsersCurrentHost    (Console) # Global modification of console UI/colors
# $Profile.CurrentUserCurrentHost (Console) # Noah's user-specific console settings (none)
#		   CurrentUserCurrentHost (ISE)		# Does not exist
# 		   AllUsersCurrentHost 	  (ISE)     # ISE Tabs script

# List the profiles using:
# $PROFILE | Format-List -Force

# AllUsersAllHosts       (All):     C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1
# CurrentUserAllHosts    (All):     C:\Users\Noah\Documents\WindowsPowerShell\profile.ps1
# AllUsersCurrentHost    (Console): C:\Windows\System32\WindowsPowerShell\v1.0\Microsoft.PowerShell_profile.ps1
# CurrentUserCurrentHost (Console): C:\Users\Noah\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
# AllUsersCurrentHost    (ISE):     C:\Windows\System32\WindowsPowerShell\v1.0\Microsoft.PowerShellISE_profile.ps1
# CurrentUserCurrentHost (ISE): 	C:\Users\Noah\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1


# **********------------------------------------------------------********** #
# |                       ~~ MODIFY HOST WINDOW ~~                         | #
# **********------------------------------------------------------********** #

# Creating a variable for easy configuration of console. 
$console = $Host.UI.RawUI

# Window title can be changed here.
$console.WindowTitle = $console.WindowTitle + '-' + $env:username

# Buffer size config.
# $buffer = $console.BufferSize
# $buffer.Width = 200
# $buffer.Height = 9001
# $console.BufferSize = $buffer

# Window size config.
# $size = $console.WindowSize
# $size.Width = 85
# $size.Height = 60
# $console.WindowSize = $size

# Powershell color config.
$Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
$Host.UI.RawUI.ForegroundColor = 'White'
# $Host.PrivateData.ErrorForegroundColor = 'Red'
# $Host.PrivateData.ErrorBackgroundColor = $bckgrnd
# $Host.PrivateData.WarningForegroundColor = 'Yellow'
# $Host.PrivateData.WarningBackgroundColor = $bckgrnd
$Host.PrivateData.DebugForegroundColor = 'DarkYellow'
$Host.PrivateData.DebugBackgroundColor = $bckgrnd
$Host.PrivateData.VerboseForegroundColor = 'Cyan'
$Host.PrivateData.VerboseBackgroundColor = $bckgrnd
$Host.PrivateData.ProgressForegroundColor = 'Magenta'
$Host.PrivateData.ProgressBackgroundColor = $bckgrnd

# Testing/debugging for colour settings.
# Write-Host ''
# Write-Host -NoNewline "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
# Write-Output "~~~~~~~~~~~~~~~~~~~~~~~~~~"
# Write-Host ''
# Write-Output "TEXT: This is a test message."
# Write-Verbose "This is a verbose message." -Verbose
# Write-Warning "This is a warning message."
# Write-Error "This is an error message."
# Write-Host -NoNewline "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
# Write-Output "~~~~~~~~~~~~~~~~~~~~~~~~~~"
# Write-Host ''

# Clearing window after color configuration is done. 
# Comment this of to enable previous debugging/testing module.
# Clear-Host


# **********------------------------------------------------------********** #
# |                   ~~ CLEAN-UP AND WELCOME MESSAGE ~~                   | #
# **********------------------------------------------------------********** #

# Clearing window after aliases are set.
# Clear-Host

Write-Host 'Windows Powershell' $PsVersionTable.PSVersion -ForegroundColor Yellow
Write-Host 'Copyright (C) 2012 Microsoft Corporation. All rights reserved.' -ForegroundColor Yellow
Write-Host 
Write-Host 'Greetings' $env:username'. It is currently' (Get-date)
Write-Host 
Write-Host

# Set starting location to "home" only if already set to System32
$loc = (Get-Location).path
if ($loc -eq 'C:\Windows\System32') {Set-Location ~} elseif ($loc -eq 'C:\Windows\System32\WindowsPowerShell\v1.0') {Set-Location ~}

# List user-defined aliases (need to be manually updated)
$loc = (Get-Location).path  # update location
if ($loc -eq 'C:\Users\Noah') {Get-Alias $Aliases}  # prevent listing in PyCharm
