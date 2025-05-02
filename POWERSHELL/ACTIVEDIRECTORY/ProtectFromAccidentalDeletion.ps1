Import-Module activedirectory 
# Path to search in for OU's
$searchbase = "DC=uson,DC=local"
# Get all the OU's that are protected
$protectedOrganizationalUnits = Get-ADOrganizationalUnit -searchbase $searchbase -filter * -Properties ProtectedFromAccidentalDeletion | where {$_.ProtectedFromAccidentalDeletion -eq $true}
# Display OU's that are protected
$protectedOrganizationalUnits | Select DistinguishedName, ProtectedFromAccidentalDeletion, Name
# Disable protection
#$protectedOrganizationalUnits | Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $false