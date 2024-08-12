Get-ADComputer -Filter {Operatingsystem -Like 'Windows*' -and Enabled -eq 'true'}  -ErrorAction SilentlyContinue -Properties *  |
    Select-Object -ExpandProperty Name | Out-File -FilePath C:\LazyWinAdmin\Servers\Domain-Computers.txt -Append

    Get-Content 'C:\LazyWinAdmin\Servers\Domain-Computers.txt' | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File 'C:\LazyWinAdmin\Servers\Domain-Computers-Alive.txt' -append
  } else { 
  write-output "$_ is Dead!!!" | Out-File 'C:\LazyWinAdmin\Servers\Domain-Computers-Dead.txt' -append}}


