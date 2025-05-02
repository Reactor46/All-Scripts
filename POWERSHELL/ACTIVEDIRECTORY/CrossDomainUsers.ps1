#-Domain cloud.local -SamAccountName etajima
#Get-ADObject -filter "objectsid -eq 'S-1-5-'" -Properties memberof  | Select-Object -ExpandProperty memberof

$UserSAM = 'etajima'
$DomainADC = 'msodc04.uson.local'
$DomainBDC = '172.20.0.5'
$DomainACreds = Get-Credential
$DomainBCreds = Get-Credential

New-PSDrive -Name DOMAINA -PSProvider ActiveDirectory -Server $DomainADC -Root "//RootDSE/" -Scope Global -Credential $DomainACreds | out-null
New-PSDrive -Name DOMAINB -PSProvider ActiveDirectory -Server $DomainBDC -Root "//RootDSE/" -Scope Global -Credential $DomainBCreds | out-null
set-location DOMAINA:
$ADUser = get-aduser $UserSAM
Set-location DOMAINB:
$LDAPFilter = '(objectSID={0})' -f $ADUser.SID.ToString()
$FSP = Get-ADObject -LDAPFilter $LDAPFilter
$LDAPFilter = '(member={0})' -f $FSP.DistinguishedName
$Groups = Get-ADGroup -LDAPFilter $LDAPFilter