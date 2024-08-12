<#
.SYNOPSIS
This script will export to a '|' delimeted file the active directory subnets belonging to the typed site.

.PARAMETER Sitename
The active Directory sitename

.PARAMETER filepath
Path to the file where we want to export the data.

.EXAMPLE
 
C:\PS>Get-ADSiteSubnets.ps1 -Sitename HQ -filePath  .\Headquarters-subnets.txt

This Example will export to a '|' delimeted file the active directory subnets belonging to the site named HQ.

.NOTES
         NAME......:  Get-ADSiteSubnets
         AUTHOR....:  Guillermo Serrano
         LAST EDIT.:  06/13/2017
         CREATED...:  06/13/2017
#>


[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$Sitename,
   [string]$filePath
)

$site = get-adreplicationsite -filter {name -eq $sitename}
if($site)
{
	$subnets = Get-ADReplicationSubnet -filter {site -eq $site.DistinguishedName} -Properties Description
}
else
{
	write-host "Unable find the typed site"
}

$subnets |export-csv -path $filepath -NoTypeInformation -Delimiter '|'