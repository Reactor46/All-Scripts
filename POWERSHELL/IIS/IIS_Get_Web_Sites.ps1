Import-Module WebAdministration

Get-Website | select name,id,state,physicalpath, 
@{n="Bindings"; e= { ($_.bindings | select -expa collection) -join ';' }} ,
@{n="LogFile";e={ $_.logfile | select -expa directory}}, 
@{n="attributes"; e={($_.attributes | % { $_.name + "=" + $_.value }) -join ';' }} |
Export-Csv -Path "\\fbv-wbdv20-d01\D$\IIS_Web_Sites_ALL.csv" -NoTypeInformation -Append #Change this path to suit your environment