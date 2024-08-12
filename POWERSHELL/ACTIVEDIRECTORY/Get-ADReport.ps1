Import-Module ActiveDirectory

Get-ADObject -Filter 'Name -like "*"' | 
Where-Object {$_.ObjectClass -eq "user" -or $_.ObjectClass -eq "computer" -or $_.ObjectClass -eq "group" -or $_.ObjectClass -eq "organizationalUnit"} |
    Sort-Object ObjectClass | Export-CSV C:\LazyWinAdmin\ExportAD.csv -notypeinformation



$Path = 'C:\LazyWinAdmin\ADUsers.csv'

Get-ADUser -Filter * |

Select-Object Name,Enabled,UserPrincipalName | Export-Csv -Path $Path –notypeinformation