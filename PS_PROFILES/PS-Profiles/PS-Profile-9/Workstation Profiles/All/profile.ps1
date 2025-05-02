# ************************************************************************** #
# ||    	              ~~ AllUsersAllHosts (All) ~~                    || #
# ||                         Written by Noah Hopkins                      || #
# ||                       Last edited: October, 2018                     || #
# ************************************************************************** #

# NOTE: This profile will affect all users, and both Powershell, and 
# Powershell ISE instances.

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
# |                          ~~ IMPORT MODULES ~~                          | #
# **********------------------------------------------------------********** #

Import-Module posh-git


# **********------------------------------------------------------********** #
# |                          ~~ GLOBAL ALIASES ~~                          | #
# **********------------------------------------------------------********** #

# Creating an alias for notepad++.
# [np [file name]] opens file in notepad++. Not specifying [file name] just 
# opens notepad++.
New-Item alias:np -value "C:\Program Files (x86)\Notepad++\notepad++.exe"

# Alias for easier access to "repos" directory.
Function Open-Repos {Set-Location "C:\Users\Noah\repos"}
New-Item alias:repos -value Open-Repos

## Alias for easier access to "scripting" directory.
Function Open-ScriptingLoc {Set-Location "D:\Noah\OneDrive\Scripting"}
New-Item alias:scripts -value Open-ScriptingLoc

# Alias for easier access to "PycharmProjects" directory.
Function Open-PyCharmProjects {Set-Location "C:\Users\Noah\PycharmProjects"}
New-Item alias:pych -value Open-PyCharmProjects

# Alias for easier access to "CS project PycharmProject" directory.
Function Open-PyCharmCSProject {Set-Location "C:\Users\Noah\PycharmProjects\rna_cs_project"}
New-Item alias:csp -value Open-PyCharmCSProject

# Alias for hibernation.
Function Hibernate-Computer {C:\Windows\System32\rundll32.exe powrprof.dll,SetSuspendState Hibernate}
New-Item alias:hibernate -value Hibernate-Computer

# Function Sleep-Computer {C:\Windows\System32\rundll32.exe powrprof.dll,SetSuspendState 0,1,0}
# New-Item alias:sleepc -value Sleep-Computer

# ## Alias for easier access to Noah's Onedrive directory.
Function Open-OnedriveLoc {Set-Location "D:\Noah\OneDrive\"}
New-Item alias:od -value Open-OnedriveLoc

## PS Drives
# New-PSDrive -Name od -PSProvider FileSystem -Root "C:\Users\Noah\OneDrive\"
# New-PSDrive -Name Scripting -PSProvider FileSystem -Root "C:\Users\Noah\OneDrive\Scripting"

## Alias for easier acces to Noah's python profile.
Function Explore-PyProfileDir {Invoke-Item "C:\Users\Noah\AppData\Roaming\Python\Python27\site-packages"}
New-Item alias:pydir -value Explore-PyProfileDir

$Aliases = 'np', 'repos', 'scripts', 'pych', 'csp', 'hibernate', 'od', 'pydir'

# Clear host after aliases have been set
Clear-Host