#### Massive Server List
#### Let's delete the existing results.
$FileNamePath = "C:\LazyWinAdmin\Servers\RESULTS"
$FileName = "$FileNamePath\CREDITONEAPP.TST\C1A.TST.Alive.txt","$FileNamePath\CREDITONEAPP.TST\C1A.TST.txt","$FileNamePath\CREDITONEAPP.TST\C1A.TST.Dead.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.Dead.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.txt","$FileNamePath\FNBM.CORP\FNBM.txt","$FileNamePath\FNBM.CORP\FNBM.Alive.txt","$FileNamePath\FNBM.CORP\FNBM.Dead.txt","$FileNamePath\PHX.FNBM.CORP\PHX.txt","$FileNamePath\PHX.FNBM.CORP\PHX.Alive.txt","$FileNamePath\PHX.FNBM.CORP\PHX.Dead.txt"
if (Test-Path $FileName -ErrorAction SilentlyContinue) {
  Remove-Item $FileName -ErrorAction SilentlyContinue
}
$ResultsPath = "C:\LazyWinAdmin\Servers\RESULTS"

#### FNBM.CORP
Get-ADComputer -Server LASDC02.FNBM.CORP -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.txt"  -Append

Get-Content "$ResultsPath\FNBM.CORP\FNBM.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.Dead.txt" -Append }}

#### PHX.FNBM.CORP
Get-ADComputer -Server PHXDC03.PHX.FNBM.CORP -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.txt"  -Append

Get-Content "$ResultsPath\PHX.FNBM.CORP\PHX.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.Dead.txt" -Append }}
  
#### C1B.BIZ
  Get-ADComputer -Server LASAUTH01.CREDITONEAPP.BIZ -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.txt"  -Append

Get-Content "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Dead.txt" -Append }}

#### C1B.TST
  Get-ADComputer -Server LASAUTHTST01.CREDITONEAPP.TST -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.txt"  -Append

Get-Content "$ResultsPath\CREDITONEAPP.TST\C1A.TST.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.TST.Dead.txt" -Append }}


Copy-Item "$ResultsPath\CREDITONEAPP.TST\C1A.TST.Alive.txt" -Destination "$ResultsPath\Alive\C1A.TST.Alive.txt"
Copy-Item "$ResultsPath\FNBM.CORP\FNBM.Alive.txt" -Destination "$ResultsPath\Alive\FNBM.Alive.txt"
Copy-Item "$ResultsPath\PHX.FNBM.CORP\PHX.Alive.txt" -Destination "$ResultsPath\Alive\PHX.Alive.txt"
Copy-Item "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.txt" -Destination "$ResultsPath\Alive\C1A.BIZ.Alive.txt"
Get-Content "$ResultsPath\Alive\C1A.TST.Alive.txt","$ResultsPath\Alive\FNBM.Alive.txt","$ResultsPath\Alive\PHX.Alive.txt","$ResultsPath\Alive\C1A.BIZ.Alive.txt" |
    Group | where {$_.count -eq 1} | % {$_.group[0]} | Set-Content "$ResultsPath\Alive\ALL.Alive.txt"
    Remove-Item "$ResultsPath\Alive\C1A.TST.Alive.txt","$ResultsPath\Alive\FNBM.Alive.txt","$ResultsPath\Alive\PHX.Alive.txt","$ResultsPath\Alive\C1A.BIZ.Alive.txt"



#$ChkLicStatus = Get-Content "$ResultsPath\Alive\ALL.Alive.txt"

#.\Get-WindowsLicenseDetails.ps1 -TextFile $ChkLicStatus | Select PSComputerName, LicenseStatus, Description, LicenseFamily, PartialProductKey