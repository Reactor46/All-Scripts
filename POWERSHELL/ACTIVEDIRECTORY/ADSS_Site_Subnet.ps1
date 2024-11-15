<################################################################################################################################
 Title:   Script to acquire the Site Name and associated subnets.
 Preface: Script to acquire the Site Name and associated subnets. This script will return SiteName and Subnet and export it ADSS.csv

 Author:  Vikram Bedi
 Website: www.vikrambedi.com
 Blog:    http://www.vikrambedi.com/community/
 Email:   vikram.bedi.it@gmail.com  
 Powershell Version: Powershell v2.0  
 Script Version: v1.0
 
 Change Log:
 v1.0 - Script to acquire subnets and associated subnets and return SiteName and subnet to csv file.
################################################################################################################################>

$sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
 
$sitesubnets = @()
 
foreach ($site in $sites)
{
	foreach ($subnet in $site.subnets){
	   $temp = New-Object PSCustomObject -Property @{
	   'SiteName' = $site.Name
	   'Subnet' = $subnet; }
	    $sitesubnets += $temp
	}
}
 
$sitesubnets
$sitesubnets | Export-Csv $PSScriptRoot\ADSS.csv