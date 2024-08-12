# Exporting DL Objects
Get-DistributionGroupMember -Identity  "<DL-NAME>" | where {$_.RecipientType -eq 'UserMailbox'} |Select-Object name, WindowsLiveID |  Export-Csv ".\filepath\<DL-NAME>.csv"
#Export Only DL Names from the Parent Group
Get-DistributionGroupMember -Identity  "<DL-NAME>" | where {$_.RecipientType -ne 'UserMailbox'} | Select-Object name | Export-Csv  ".\filepath\<DL-NAME_dl>.csv"

# Exporting Nested DL Objects With individual File 

Import-Csv  ".\filepath\<DL-NAME_dl>.csv" |

foreach{

Get-DistributionGroupMember -Identity   $_.Name  | Select-Object name, WindowsLiveID | Export-Csv -Path  .\filepath\"$($_.Name)".csv -NoTypeInformation

}

