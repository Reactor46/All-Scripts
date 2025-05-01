# ************************************************************************** #
# ||    	          ~~ CurrentUserCurrentHost (Console) ~~              || #
# ||                         Written by Noah Hopkins                      || #
# ||                       Last edited: October, 2018                     || #
# ************************************************************************** #

# NOTE: This profile will only affect Noah's Powershell instances, but not 
# Powershell ISE (I think).

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
# |               ~~ IMPORT MODULES / SET ENV VARIABLES ~~                 | #
# **********------------------------------------------------------********** #

## Importing the Powershell Community Extension module.
# TODO: Need to fix!

# Import-Module PSCX

# Import-Module posh-git

## Posh-Git enviroment fix. Don't know if nescessary.
## From http://www.imtraum.com/blog/streamline-git-with-powershell/.
# $env:path += ";" + (Get-Item "Env:ProgramFiles(x86)").Value + "\Git\bin"

## Load Posh-Git profile.
# . 'C:\Users\Noah\documents\Projects\posh-git\profile.noah.ps1'

## Related to Posh-Git activation. Found this from here: 
## https://git-scm.com/book/uz/v2/Git-in-Other-Environments-Git-in-Powershell
# . (Resolve-Path "$env:LOCALAPPDATA\GitHub\shell.ps1")
# . $env:github_posh_git\profile.example.ps1

## Related to loading Posh-Git. Found here: 
## http://stackoverflow.com/questions/12504649/how-to-use-posh-git-that-comes-with-github-for-windows-from-custom-shell
## If Posh-Git environment is defined, load it.
# if (test-path env:posh_git) {
    # . $env:posh_git
# }


# # This makes it possible to interpret python scripts without using 
# # the [python [file name]] command.
# $env:PATHEXT += ";.py"


# **********------------------------------------------------------********** #
# |                       ~~ MODIFY HOST WINDOW ~~                         | #
# **********------------------------------------------------------********** #

# NOTE: Moved to AllUsersCurrentHost (Console).


# **********------------------------------------------------------********** #
# |                              ~~ ALIASES ~~                             | #
# **********------------------------------------------------------********** #

# NOTE: Mostly moved to AllUsersAllHosts

# # Creating an alias for notepad++.
# # [np [file name]] opens file in notepad++. Not specifying [file name] just 
# # opens notepad++.
# # New-Item alias:np -value "C:\Program Files (x86)\Notepad++\notepad++.exe"
# Function Open-Npp {"C:\Program Files (x86)\Notepad++\notepad++.exe"}
# Set-Alias np "C:\Program Files (x86)\Notepad++\notepad++.exe"
# # Set-Alias alias:np -value Open-Npp

# # Alias for easier access to "repos" directory.
# Function Open-Repos {Set-Location C:\Users\Noah\repos}
# New-Item alias:repos -value Open-Repos

# ## Alias for easier access to "scripting" directory.
# Function Set-ScriptingLoc {cd "C:"
# cd "C:\Users\Noah\OneDrive\Scripting"}
# Set-alias alias:scripts Set-ScriptingLoc

# # Alias for easier access to "PycharmProjects" directory.
# Function Open-PyCharmProjects {Set-Location C:\Users\Noah\PycharmProjects}
# New-Item alias:pych -value Open-PyCharmProjects

# # Alias for easier access to "CS project PycharmProject" directory.
# Function Open-PyCharmCSProject {Set-Location C:\Users\Noah\PycharmProjects\rna_cs_project}
# New-Item alias:csp -value Open-PyCharmCSProject

# # Alias for hibernation.
# Function Hibernate-Computer {C:\Windows\System32\rundll32.exe powrprof.dll,SetSuspendState Hibernate}
# New-Item alias:hibernate -value Hibernate-Computer

# # Function Sleep-Computer {C:\Windows\System32\rundll32.exe powrprof.dll,SetSuspendState 0,1,0}
# # New-Item alias:sleepc -value Sleep-Computer

# # ## Alias for easier access to Onedrive directory.
# Function Set-OnedriveLoc {cd "C:\Users\Noah\OneDrive\"}
# Set-alias alias:od Set-OnedriveLoc

# ## PS Drives
# # New-PSDrive -Name od -PSProvider FileSystem -Root "C:\Users\Noah\OneDrive\"
# # New-PSDrive -Name Scripting -PSProvider FileSystem -Root "C:\Users\Noah\OneDrive\Scripting"

# ## Alias for easier acces to python profile.
# Function Open-PyProfileDir {Invoke-Item C:\Users\Noah\AppData\Roaming\Python\Python27\site-packages\}
# Set-Alias alias:pydir Open-PyProfileDir


# **********------------------------------------------------------********** #
# |                   ~~ CLEAN-UP AND WELCOME MESSAGE ~~                   | #
# **********------------------------------------------------------********** #

# Note: Made user-agnostic and moved to AllUsersCurrentHost.

# # Clearing window after aliases are set.
# # Clear-Host

# Write-Host 
# Write-Host 'Windows Powershell' $PsVersionTable.PSVersion -ForegroundColor Yellow
# Write-Host 'Copyright (C) 2012 Microsoft Corporation. All rights reserved.' -ForegroundColor Yellow
# Write-Host 
# Write-Host 'Greetings' $env:username'. It is currently' (Get-date)
# Write-Host 
# Write-Host

# # Setting starting location.
# $loc = (Get-Location).path
# # Write-Host $loc
# if ($loc -eq 'C:\Windows\System32') {Set-Location ~}

