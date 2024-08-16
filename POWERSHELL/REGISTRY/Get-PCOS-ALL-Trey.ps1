<#
#### Massive Server List
#### Let's delete the existing results.
$FileNamePath = "C:\LazyWinAdmin\ReadRegistry\RESULTS"
$FileName = "$FileNamePath\C1A.TST.Alive.txt","$FileNamePath\C1A.TST.txt","$FileNamePath\C1A.TST.Dead.txt","$FileNamePath\C1A.BIZ.Alive.txt","$FileNamePath\C1A.BIZ.Dead.txt","$FileNamePath\C1A.BIZ.txt","$FileNamePath\FNBM.txt","$FileNamePath\FNBM.Alive.txt","$FileNamePath\FNBM.Dead.txt","$FileNamePath\PHX.txt","$FileNamePath\PHX.Alive.txt","$FileNamePath\PHX.Dead.txt"
if (Test-Path $FileName) {
  Remove-Item $FileName
}
$ResultsPath = "C:\LazyWinAdmin\ReadRegistry\RESULTS"

#### FNBM.CORP
Get-ADComputer -Server LASDC02.FNBM.CORP -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom "$ResultsPath\FNBM.txt"  -Append

Get-Content "$ResultsPath\FNBM.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.Dead.txt" -Append }}

#### PHX.FNBM.CORP
Get-ADComputer -Server PHXDC03.PHX.FNBM.CORP -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom "$ResultsPath\PHX.txt"  -Append

Get-Content "$ResultsPath\PHX.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.Dead.txt" -Append }}
  
#### C1B.BIZ
  Get-ADComputer -Server LASAUTH01.CREDITONEAPP.BIZ -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom "$ResultsPath\C1A.BIZ.txt"  -Append

Get-Content "$ResultsPath\C1A.BIZ.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\C1A.BIZ.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\C1A.BIZ.Dead.txt" -Append }}

#### C1B.TST
  Get-ADComputer -Server LASAUTHTST01.CREDITONEAPP.TST -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom "$ResultsPath\C1A.TST.txt"  -Append

Get-Content "$ResultsPath\C1A.TST.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\C1A.TST.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\C1A.TST.Dead.txt" -Append }}


$Servers = GC "$ResultsPath\C1A.TST.Alive.txt"

ForEach($srv in $Servers){
Get-RegValue -ComputerName $srv -Hive CurrentUser -Key "Software\Microsoft\Windows\CurrentVersion\RunOnce\Registry Driver" |
    Select ComputerName, Hive, Key, Value, Data, Type |
        Export-CSV "$ResultsPath\C1A.TST-HKCU.csv" -NoTypeInformation -Append
}

$Servers = GC "$ResultsPath\FNBM.Alive.tx"

ForEach($srv in $Servers){
Get-RegValue -ComputerName $srv -Hive CurrentUser -Key "Software\Microsoft\Windows\CurrentVersion\RunOnce\Registry Driver" |
    Select ComputerName, Hive, Key, Value, Data, Type |
        Export-CSV "$ResultsPath\FNBM-HKCU.csv" -NoTypeInformation -Append
}

$Servers = GC "$ResultsPath\PHX.Alive.txt"

ForEach($srv in $Servers){
Get-RegValue -ComputerName $srv -Hive CurrentUser -Key "Software\Microsoft\Windows\CurrentVersion\RunOnce\Registry Driver" |
    Select ComputerName, Hive, Key, Value, Data, Type |
        Export-CSV "$ResultsPath\PHX-HKCU.csv" -NoTypeInformation -Append
}
#>
$Servers = GC .\RESULTS\Trey.txt

ForEach($srv in $Servers){
Get-RegValue -ComputerName $srv -Hive CurrentUser -Key "Software\Microsoft\Windows\CurrentVersion\RunOnce\" -ErrorAction SilentlyContinue |
    Select ComputerName, Hive, Key, Value, Data, Type | Export-CSV .\RESULTS\All-HKCU.csv -NoTypeInformation -Append
}