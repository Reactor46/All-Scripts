#-------------------------------------------------------------------------
# Author      : Alexandre VIOT alexandreviot.net
# FileName    : New-OU.ps1
# Version     : 1.0
# Revision    :
# Created     : 26.04.15
# Description : Powershell script creates OU into Active Directory from a CSV File.
# Remarks     : CSV file must contains Name and Path.
#
#-------------------------------------------------------------------------
param([parameter(Mandatory=$true)] [String]$FileCSV)
$listOU=Import-CSV $FileCSV -Delimiter ";"
ForEach($OU in $listOU){
 
try{
#Get Name and Path from the source file
$OUName = $OU.Name
$OUPath = $OU.Path
 
#Display the name and path of the new OU
Write-Host -Foregroundcolor Yellow $OUName $OUPath
 
#Create OU
New-ADOrganizationalUnit -Name "$OUName" -Path "$OUPath"
 
#Display confirmation
Write-Host -ForegroundColor Green "OU $OUName created"
}catch{
 
Write-Host $error[0].Exception.Message
}
 
}