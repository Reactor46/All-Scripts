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

(Get-SPSite -Limit All |  
select -ExpandProperty AllWebs |  
select -ExpandProperty AllUsers |  
?{$_.IsDomainGroup -ne $true} |  
select -Unique LoginName).count