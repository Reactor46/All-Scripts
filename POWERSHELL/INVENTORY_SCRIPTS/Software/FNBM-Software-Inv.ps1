#### Massive Server List
#### Let's delete the existing results.
$GetSoftware = "C:\LazyWinAdmin\Inventory\Software\Get-SoftwareInventory.ps1"
$FileNamePath = "C:\LazyWinAdmin\Inventory\Software\RESULTS"
$FileName = "$FileNamePath\CREDITONEAPP.TST\C1A.TST.Alive.txt","$FileNamePath\CREDITONEAPP.TST\C1A.TST.txt","$FileNamePath\CREDITONEAPP.TST\C1A.TST.Dead.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.Dead.txt","$FileNamePath\CREDITONEAPP.BIZ\C1A.BIZ.txt","$FileNamePath\Contoso.CORP\Contoso.txt","$FileNamePath\Contoso.CORP\Contoso.Alive.txt","$FileNamePath\Contoso.CORP\Contoso.Dead.txt","$FileNamePath\PHX.Contoso.CORP\PHX.txt","$FileNamePath\PHX.Contoso.CORP\PHX.Alive.txt","$FileNamePath\PHX.Contoso.CORP\PHX.Dead.txt"
if (Test-Path $FileName -ErrorAction SilentlyContinue) {
  Remove-Item $FileName -ErrorAction SilentlyContinue
}
$ResultsPath = "C:\LazyWinAdmin\Inventory\Software\RESULTS"

#### Contoso.CORP
Get-ADComputer -Server LASDC02.Contoso.CORP -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\Contoso.CORP\Contoso.txt"  -Append

Get-Content "$ResultsPath\Contoso.CORP\Contoso.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\Contoso.CORP\Contoso.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\Contoso.CORP\Contoso.Dead.txt" -Append }}

#### PHX.Contoso.CORP
Get-ADComputer -Server PHXDC03.PHX.Contoso.CORP -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\PHX.Contoso.CORP\PHX.txt"  -Append

Get-Content "$ResultsPath\PHX.Contoso.CORP\PHX.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.Contoso.CORP\PHX.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.Contoso.CORP\PHX.Dead.txt" -Append }}
  
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
Copy-Item "$ResultsPath\Contoso.CORP\Contoso.Alive.txt" -Destination "$ResultsPath\Alive\Contoso.Alive.txt"
Copy-Item "$ResultsPath\PHX.Contoso.CORP\PHX.Alive.txt" -Destination "$ResultsPath\Alive\PHX.Alive.txt"
Copy-Item "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.txt" -Destination "$ResultsPath\Alive\C1A.BIZ.Alive.txt"
Get-Content "$ResultsPath\Alive\C1A.TST.Alive.txt","$ResultsPath\Alive\Contoso.Alive.txt","$ResultsPath\Alive\PHX.Alive.txt","$ResultsPath\Alive\C1A.BIZ.Alive.txt" |
    Group | where {$_.count -eq 1} | % {$_.group[0]} | Set-Content "$ResultsPath\Alive\ALL.Alive.txt"
    Remove-Item "$ResultsPath\Alive\C1A.TST.Alive.txt","$ResultsPath\Alive\Contoso.Alive.txt","$ResultsPath\Alive\PHX.Alive.txt","$ResultsPath\Alive\C1A.BIZ.Alive.txt"

    
    

    $ServerInv = Get-Content "$ResultsPath\Alive\ALL.Alive.txt"

    ForEach ($srv in $ServerInv){
    
  .\Get-SoftwareInventory.ps1 -RemoteComputer $srv
  }
