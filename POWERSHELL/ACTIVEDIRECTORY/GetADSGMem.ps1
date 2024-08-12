<#
.SYNOPSIS
	Get AD Group membership listing and whether the accounts are active.

.DESCRIPTION
	Using Get-ADGroupMember and Get-ADUser and passing in a Security Group
	name, get the list of users and account enabled status.

.PARAMETER sg
	Security Group name (required).  Be sure to enclose the SGName if there
	are spaces in the name.

.EXAMPLE
	A sample command that uses the function or script, optionally followed
	by sample output and a description. Repeat this keyword for each example.

.NOTES
	Author: Clint Simmons
	Last Updated: 2015/06/05

.LINK
	The name of a related topic. Repeat this keyword for each related topic.

	This content appears in the Related Links section of the Help topic.

	The Link keyword content can also include a Uniform Resource Identifier
	(URI) to an online version of the same Help topic. The online version 
	opens when you use the Online parameter of Get-Help. The URI must begin
	with "http" or "https".
#>

Param(
	[Parameter(Position=0,
		Mandatory=$True,
		ValueFromPipeline=$True)]
	[string]$sg
	)

[string[]]$sname = (Get-ADGroupMember -Identity $sg -Recursive | Select-Object SAMAccountName).SAMAccountName
foreach ($i in $sname)
{
    Get-ADUser -Identity $i | Select-Object SurName, GivenName, Name, Enabled
}
