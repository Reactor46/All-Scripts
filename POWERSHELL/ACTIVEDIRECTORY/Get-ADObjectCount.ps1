
Get-ADObject -Filter {objectCategory -eq 'CN=Organizational-Unit,CN=Schema,CN=Configuration,DC=contoso,DC=com'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object
Get-ADObject -Filter {objectCategory -eq 'CN=Computer,CN=Schema,CN=Configuration,DC=contoso,DC=com'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object
Get-ADObject -Filter {objectCategory -eq 'CN=Group,CN=Schema,CN=Configuration,DC=contoso,DC=com'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object
Get-ADObject -Filter {objectCategory -eq 'CN=Person,CN=Schema,CN=Configuration,DC=contoso,DC=com'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object
Get-ADObject -Filter {objectCategory -eq 'CN=RRAS-Administration-Connection-Point,CN=Schema,CN=Configuration,DC=contoso,DC=com'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object
Get-ADObject -Filter {objectCategory -eq 'CN=Service-Connection-Point,CN=Schema,CN=Configuration,DC=contoso,DC=com'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object
Get-ADObject -Filter {objectCategory -eq 'CN=MSMQ-Configuration,CN=Schema,CN=Configuration,DC=contoso,DC=com'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object
Get-ADObject -Filter {objectCategory -eq 'CN=Intellimirror-SCP,CN=Schema,CN=Configuration,DC=contoso,DC=com'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object
Get-ADObject -Filter {name -like '*'} -SearchBase "OU=Servers,OU=Las_Vegas,DC=contoso,DC=com" -ResultSetSize $null | Measure-Object