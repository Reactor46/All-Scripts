
Get-ADComputer -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -ErrorAction SilentlyContinue -Properties *  |
    Select -ExpandProperty Name | Out-File -Encoding ascii 'C:\LazyWinAdmin\Health Scripts\Reports\Servers.txt' -Append

Get-Content 'C:\LazyWinAdmin\Health Scripts\Reports\Servers.txt' | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File -Encoding ascii 'C:\LazyWinAdmin\Health Scripts\Reports\ServersList.txt' -Append
  } else { 
  write-output "$_ is Dead!!!" | Out-File -Encoding ascii 'C:\LazyWinAdmin\Health Scripts\Reports\ServersDead.txt' -Append}}