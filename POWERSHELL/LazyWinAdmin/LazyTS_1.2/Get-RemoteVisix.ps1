CD "C:\LazyWinAdmin\LazyWinAdmin\LazyTS_1.2"


Get-ADComputer -Server LASDC02.FNBM.CORP -Filter {Operatingsystem -Like 'Windows Embedded*' -and Enabled -eq 'true'} -SearchBase "OU=Digital Displays,DC=fnbm,DC=corp" -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom ".\configs\visixPCs_ALL.txt"
   

Get-Content ".\configs\visixPCs_ALL.txt" |
    ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom ".\configs\visixPCs_Alive.txt" -Append
        } else { 
    write-output "$_" | Out-FileUtf8NoBom ".\configs\visixPCs_Dead.txt" -Append }}