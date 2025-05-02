#$Exclude = 'HealthMailbox*','Webex*','Public*','DEVTEST*','EXCH*','750*','*Service*','Training*','Conference*','*Calendar*','*test*','*em7*','WINSYS*'

#$filter = ($Exclude | foreach {'(Name -ne ' + "'$_')"}) -join ' -and '

Get-ADUser -Filter {(Enabled -eq $true) -and (mail -like "*") -and (ObjectCategory -eq "CN=Person,CN=Schema,CN=Configuration,DC=fnbm,DC=corp") -and (Department -like "*")} -SearchBase 'DC=FNBM,DC=CORP' -Properties * | 
   Select Name   , EmailAddress   , Department   , Description | Export-CSV -Path C:\LazyWinAdmin\Reports\AD-Users.csv -NoTypeInformation -Append
     