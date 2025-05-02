<#
.SYNOPSIS
	Number of unique user in SharePoint farm.

.DESCRIPTION
	The script can count the all unique user in a SharePoint farm. 

.NOTES
	Author: Tibor Revesz
	Requires: PowerShell v2
	Version: 1.0
 
.EXAMPLE
	.\Get-AllSPUser.ps1
	Runs the script from the current directory.
#>

# Check and add the Microsoft.SharePoint.PowerShell snapin if it is not loaded.
If((Get-PSSnapin Microsoft.SharePoint.PowerShell –EA SilentlyContinue) –eq
$null){Add-PSSnapin Microsoft.SharePoint.PowerShell}

# Get all site collection and the included subsites.
Get-SPSite -Limit ALL |
Get-SPWeb -Limit ALL |
# Get all users from a subsite.
%{Get-SPUser -Web $_.Url -Limit ALL} |
# Check if the object is not a group.
Where-Object {$_.IsDomainGroup -ne $true} |
# Select all unique accounts.
select -Unique loginname |
# Just count the all objects.
Measure-Object |
select Count