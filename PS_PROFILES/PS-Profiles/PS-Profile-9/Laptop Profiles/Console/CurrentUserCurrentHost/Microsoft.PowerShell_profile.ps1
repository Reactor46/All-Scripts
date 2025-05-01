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

# Moved to AllUsersAllHosts
# Import-Module posh-git
