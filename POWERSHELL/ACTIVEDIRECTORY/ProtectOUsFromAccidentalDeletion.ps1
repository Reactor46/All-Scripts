Get-ADOrganizationalUnit -Filter * -Properties *|
Where-Object -Property ProtectedFromAccidentalDeletion -EQ $false|
Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $true