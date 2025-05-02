#### Massive Server List
#### Let's delete the existing results.
$FileNamePath = "C:\LazyWinAdmin\Servers\RESULTS"
$FileName = "$FileNamePath\CREDITONEAPP.TST\C1A.TST.DPC.Alive.txt","$FileNamePath\CREDITONEAPP.TST\C1A.TST.txt","$FileNamePath\CREDITONEAPP.TST\C1A.TST.DPC.Dead.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.DPC.Alive.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.DPC.Dead.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.txt","$FileNamePath\FNBM.CORP\FNBM.txt","$FileNamePath\FNBM.CORP\FNBM.DPC.Alive.txt","$FileNamePath\FNBM.CORP\FNBM.DPC.Dead.txt","$FileNamePath\PHX.FNBM.CORP\PHX.txt","$FileNamePath\PHX.FNBM.CORP\PHX.DPC.Alive.txt","$FileNamePath\PHX.FNBM.CORP\PHX.DPC.Dead.txt"
if (Test-Path $FileName -ErrorAction SilentlyContinue) {
  Remove-Item $FileName -ErrorAction SilentlyContinue
}
$ResultsPath = "C:\LazyWinAdmin\Servers\RESULTS"

#### FNBM.CORP
Get-ADComputer -Server LASDC02.FNBM.CORP -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.DPC.txt"  -Append

Get-Content "$ResultsPath\FNBM.CORP\FNBM.DPC.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.DPC.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.DPC.Dead.txt" -Append }}

#### PHX.FNBM.CORP
Get-ADComputer -Server PHXDC03.PHX.FNBM.CORP -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.DPC.txt"  -Append

Get-Content "$ResultsPath\PHX.FNBM.CORP\PHX.DPC.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.DPC.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.DPC.Dead.txt" -Append }}
  
#### C1B.BIZ
  Get-ADComputer -Server LASAUTH01.CREDITONEAPP.BIZ -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.DPC.txt"  -Append

Get-Content "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.DPC.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.DPC.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.DPC.Dead.txt" -Append }}

#### C1B.TST
  Get-ADComputer -Server LASAUTHTST01.CREDITONEAPP.TST -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.DPC.txt"  -Append

Get-Content "$ResultsPath\CREDITONEAPP.TST\C1A.TST.DPC.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.DPC.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.TST.DPC.Dead.txt" -Append }}


Copy-Item "$ResultsPath\CREDITONEAPP.TST\C1A.TST.DPC.Alive.txt" -Destination "$ResultsPath\Alive\C1A.TST.DPC.Alive.txt"
Copy-Item "$ResultsPath\FNBM.CORP\FNBM.DPC.Alive.txt" -Destination "$ResultsPath\Alive\FNBM.DPC.Alive.txt"
Copy-Item "$ResultsPath\PHX.FNBM.CORP\PHX.DPC.Alive.txt" -Destination "$ResultsPath\Alive\PHX.DPC.Alive.txt"
Copy-Item "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.DPC.Alive.txt" -Destination "$ResultsPath\Alive\C1A.BIZ.DPC.Alive.txt"
Get-Content "$ResultsPath\Alive\C1A.TST.DPC.Alive.txt","$ResultsPath\Alive\FNBM.DPC.Alive.txt","$ResultsPath\Alive\PHX.DPC.Alive.txt","$ResultsPath\Alive\C1A.BIZ.DPC.Alive.txt" |
    Group | where {$_.count -eq 1} | % {$_.group[0]} | Set-Content "$ResultsPath\Alive\ALL.DPC.Alive.txt"
    Remove-Item "$ResultsPath\Alive\C1A.TST.DPC.Alive.txt","$ResultsPath\Alive\FNBM.DPC.Alive.txt","$ResultsPath\Alive\PHX.DPC.Alive.txt","$ResultsPath\Alive\C1A.BIZ.DPC.Alive.txt"



$DPC_LIST = Get-Content "$ResultsPath\Alive\ALL.DPC.Alive.txt"

ForEach($dpc in $DPC_LIST){
    .\Get-SoftwareTitle.ps1 -ComputerName $dpc -Title "Infognition ScreenPressor v2.1*" | Select ComputerName,Title, UninstallString | Export-Csv .\DesktopsWithNetwrix.csv -Append -NoTypeInformation
    .\Get-SoftwareTitle.ps1 -ComputerName $dpc -Title "Netwrix*" | Select ComputerName,Title, UninstallString | Export-Csv .\DesktopsWithNetwrix.csv -Append -NoTypeInformation
    }