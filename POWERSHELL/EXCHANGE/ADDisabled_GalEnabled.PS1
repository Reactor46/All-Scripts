

# ========================================================
#
# 	Script Information
#
#	Title:          find users disabled in AD but not hidden from the GAL
#                   I have commented out the last part of this script. This will
#                   disable the listed users if it is un-commented. To uncomment
#                   remove the "#" before the last pipe.
#	Author:			steve.lindsey
#	Originally created:	5/14/14 
#	
#	Description:	updates UPN suffix of users in a specific OU
#	
# ========================================================
Get-Mailbox -Filter{(HiddenFromAddressListsEnabled -eq $false) -AND (UserAccountControl -eq "AccountDisabled, NormalAccount")} #| set-mailbox -HiddenFromAddressListsEnabled $true