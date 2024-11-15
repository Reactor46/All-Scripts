#change (-Identity "USON VPN USERS") to relative group..
Import-Module ActiveDirectory
Get-ADGroupMember -Identity "USON VPN USERS" -Recursive | %{Get-ADUser -Identity $_.distinguishedName -Properties Enabled | ?{$_.Enabled -eq $false}} | Select DistinguishedName,Enabled | Export-Csv c:\scripts\outfile\disabledusers.csv -NoTypeInformation