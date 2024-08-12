Enable-ADOptionalFeature 'Recycle Bin Feature' -Scope ForestOrConfigurationSet -Target 'Contoso.CORP' -server LASDC10.Contoso.CORP -WhatIf
Get-ADOrganizationalUnit -Filter * -Properties * | where {$_.ProtectedFromAccidentalDeletion -eq $false} | Export-CSV “C:\LazyWinAdmin\Servers\OUStatus.CSV” -NoTypeInformation -Append 
Get-ADObject -filter {(ObjectClass -eq 'user')} | Set-ADObject -ProtectedFromAccidentalDeletion:$true -WhatIf
Get-ADOrganizationalUnit -filter * | Set-ADObject -ProtectedFromAccidentalDeletion:$true -WhatIf


<#
$AllOUs = Get-ADOrganizationalUnit –Property Identity
ForEach ($ThisOU in $AllOUs)
{
Set-ADOrganizationalUnit -Identity $ThisOU -ProtectedFromAccidentalDeletion $True -WhatIf
}

$OUProtectReport = "C:LazyWinAdmin\Servers\OUProtectionReport.CSV"
Remove-item $OUProtectReport -ErrorAction SilentlyContinue
$ThisDomain = “Contoso.CORP”
$ThisStr = “OU Protection Setting Status in Active Directory Domain:”+$ThisDomain
Add-Content $OUProtectReport $ThisSTR
$ThisSTR = "OU Name, Is Protection Enabled?"
Add-Content $OUProtectReport $ThisSTR
$AllOUs = Get-ADOrganizationalUnit –Property Identity
ForEach ($ThisOU in $AllOUs)
{
Set-ADOrganizationalUnit –Server $ThisDomain -Identity $ThisOU -ProtectedFromAccidentalDeletion $True
$ProtStatus = Get-ADOrgaizationalUnit –Identity $ThisOU –Properties ProtectedFromAccidentalDeletion
$ProtOrNot = "No"
IF ($ProtStatus.ProtectedFromAccidentalDeletion -eq $True)
{
$ProtOrNot = "Yes"
}
$ThisSTR = $ThisOU+","+$ProtOrNot
Add-Content $OUProtectReport $ThisSTR
}
Write-Host "Protection Settings have been modified on the Organizational Units. Please check C:LazyWinAdmin\Servers\OUProtectionReport.CSV"

#>
