cls
$csv="S:\IT\SupportServices\NOC\Scripts\Expiring_Users\ExpiringAccounts.csv"
$ErrorActionPreference = "SilentlyContinue" 
Clear-Content $csv -Force
$counter= 0
$disabledaccnt = 0x02

## $list=Search-ADAccount -AccountExpiring -TimeSpan "30" -Server "Contoso.corp" -SearchBase "DC=contoso,DC=com"  | select name,accountexpirationdate,lastlogondate,samaccountname,emailaddress,distinguishedname,description
$list= Search-ADAccount -AccountExpiring -TimeSpan "30" -Server "Contoso.corp" -SearchBase "DC=contoso,DC=com"| Get-ADUser -pr manager,userAccountControl,mail,description,lastlogondate,accountexpirationdate | where {($_.userAccountControl -band $disabledaccnt) -ne $disabledaccnt} | select name,accountexpirationdate,lastlogondate,samaccountname,distinguishedname,description,manager,userAccountControl,mail
#$email = Get-ADUser -Filter {samaccountname -like "$sam"} -Server "Contoso.corp" -SearchBase "DC=contoso,DC=com" -Properties emailaddress | select -property emailaddress

foreach ($line in $list)
{
Write-Host 
#$managersam = Get-ADUser-filter { DistinguishedName -like  "$_.manager"} -pr samaccountname
##return $line.samaccountname
#$email = Get-ADUser -Filter {samaccountname -like $sam} -Server "Contoso.corp" -SearchBase "DC=contoso,DC=com" -Properties emailaddress | select -property emailaddress
#$arr = @()

$myobj = "" | Select "name","samaccountname","accountexpirationdate","emailaddress","lastlogondate","distinguishedname","description","manager","managersam","userAccountControl"
$myobj.name = $line.name
$myobj.samaccountname = $line.samaccountname
$myobj.accountexpirationdate = $line.accountexpirationdate
$myobj.emailaddress = $line.mail
$myobj.lastlogondate = $line.lastlogondate
$myobj.distinguishedname = $line.distinguishedname
$myobj.description = $line.description
$myobj.manager = $line.manager
#$myobj.managersam = Get-ADUser -filter { DistinguishedName -like  $_.manager} -pr samaccountname
$myobj.managersam = (Get-ADUser (Get-ADUser $line.samaccountname -properties manager).manager).samaccountname
$myobj.userAccountControl = $line.userAccountControl
#$arr += $myobj
#$arr | Export-Csv -Append $csv
$myobj | Export-Csv -Append $csv -NoTypeInformation
$counter++
}
Write-Host "Found $counter users"
Write-Host
Write-Host "Saved to $csv"
Write-Host