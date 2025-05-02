#  Script name:     Set AD Users to Expire.ps1
#  Created on:      04-26-2016
#  Author:          Jack Musick
#  Purpose:         Sets all AD users to expire, minus administrators.

$DomainAdminsDN = (Get-ADGroup 'Domain Admins').DistinguishedName
$AdministrationsDN = (Get-ADGroup 'Domain Admins').DistinguishedName
Get-ADUser -Filter { (memberof -ne $DomainAdminsDN) -and (memberof -ne $AdministrationsDN) }

# Actions to take against accounts. Not finished as will effect service accounts in AD that are normal, Domain Users.
# A possible fix for this would be to apply to only a specific OU.