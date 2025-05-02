#### Massive Server List
#### Let's delete the existing results.
$FileNamePath = "C:\LazyWinAdmin\Desktops\RESULTS"
$FileName = "$FileNamePath\CREDITONEAPP.TST\C1A.TST.Alive.Desktops.txt","$FileNamePath\CREDITONEAPP.TST\C1A.TST.Desktops.txt","$FileNamePath\CREDITONEAPP.TST\C1A.TST.Dead.Desktops.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.Desktops.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.Dead.Desktops.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.Desktops.txt","$FileNamePath\FNBM.CORP\FNBM.Desktops.txt","$FileNamePath\FNBM.CORP\FNBM.Alive.Desktops.txt","$FileNamePath\FNBM.CORP\FNBM.Dead.Desktops.txt","$FileNamePath\PHX.FNBM.CORP\PHX.Desktops.txt","$FileNamePath\PHX.FNBM.CORP\PHX.Alive.Desktops.txt","$FileNamePath\PHX.FNBM.CORP\PHX.Dead.Desktops.txt"
if (Test-Path $FileName -ErrorAction SilentlyContinue) {
  Remove-Item $FileName -ErrorAction SilentlyContinue
}
$ResultsPath = "C:\LazyWinAdmin\Desktops\RESULTS"

#### FNBM.CORP
Get-ADComputer -Server LASDC02.FNBM.CORP -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.Desktops.txt"  -Append

Get-Content "$ResultsPath\FNBM.CORP\FNBM.Desktops.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.Alive.Desktops.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.Dead.Desktops.txt" -Append }}

#### PHX.FNBM.CORP
Get-ADComputer -Server PHXDC03.PHX.FNBM.CORP -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.Desktops.txt"  -Append

Get-Content "$ResultsPath\PHX.FNBM.CORP\PHX.Desktops.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.Alive.Desktops.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.Dead.Desktops.txt" -Append }}
  
#### C1B.BIZ
  Get-ADComputer -Server LASAUTH01.CREDITONEAPP.BIZ -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Desktops.txt"  -Append

Get-Content "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Desktops.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.Desktops.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Dead.Desktops.txt" -Append }}

#### C1B.TST
  Get-ADComputer -Server LASAUTHTST01.CREDITONEAPP.TST -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.Desktops.txt"  -Append

Get-Content "$ResultsPath\CREDITONEAPP.TST\C1A.TST.Desktops.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.Alive.Desktops.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.TST.Dead.Desktops.txt" -Append }}


Copy-Item "$ResultsPath\CREDITONEAPP.TST\C1A.TST.Alive.Desktops.txt" -Destination "$ResultsPath\Alive\C1A.TST.Alive.Desktops.txt"
Copy-Item "$ResultsPath\FNBM.CORP\FNBM.Alive.Desktops.txt" -Destination "$ResultsPath\Alive\FNBM.Alive.Desktops.txt"
Copy-Item "$ResultsPath\PHX.FNBM.CORP\PHX.Alive.Desktops.txt" -Destination "$ResultsPath\Alive\PHX.Alive.Desktops.txt"
Copy-Item "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.Desktops.txt" -Destination "$ResultsPath\Alive\C1A.BIZ.Alive.Desktops.txt"
Get-Content "$ResultsPath\Alive\C1A.TST.Alive.Desktops.txt","$ResultsPath\Alive\FNBM.Alive.Desktops.txt","$ResultsPath\Alive\PHX.Alive.Desktops.txt","$ResultsPath\Alive\C1A.BIZ.Alive.Desktops.txt" |
    Group | where {$_.count -eq 1} | % {$_.group[0]} | Set-Content "$ResultsPath\Alive\ALL.Alive.Desktops.txt"
    Remove-Item "$ResultsPath\Alive\C1A.TST.Alive.Desktops.txt","$ResultsPath\Alive\FNBM.Alive.Desktops.txt","$ResultsPath\Alive\PHX.Alive.Desktops.txt","$ResultsPath\Alive\C1A.BIZ.Alive.Desktops.txt"
