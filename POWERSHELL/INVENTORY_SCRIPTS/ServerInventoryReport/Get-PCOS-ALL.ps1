#### Massive Server List
#### Test Paths
$PathRoot = "$PSSCriptRoot\RESULTS"
$PathTST = "$PSSCriptRoot\RESULTS\CREDITONEAPP.TST"
$PathBIZ = "$PSSCriptRoot\RESULTS\CREDITONEAPP.BIZ"
$PathFNBM = "$PSSCriptRoot\RESULTS\FNBM.CORP"
$PathPHX = "$PSSCriptRoot\RESULTS\PHX.FNBM.CORP"
If (!(Test-Path -PathType Container $PathRoot)){ New-Item -ItemType Directory -Force -Path "$PSSCriptRoot\RESULTS" }
If (!(Test-Path -PathType Container $PathTST)){ New-Item -ItemType Directory -Force -Path "$PSSCriptRoot\RESULTS\CREDITONEAPP.TST" }
If (!(Test-Path -PathType Container $PathBIZ)){ New-Item -ItemType Directory -Force -Path "$PSSCriptRoot\RESULTS\CREDITONEAPP.BIZ" }
If (!(Test-Path -PathType Container $PathFNBM)){ New-Item -ItemType Directory -Force -Path "$PSSCriptRoot\RESULTS\FNBM.CORP" }
If (!(Test-Path -PathType Container $PathPHX)){ New-Item -ItemType Directory -Force -Path "$PSSCriptRoot\RESULTS\PHX.FNBM.CORP" }
#### Let's delete the existing results.
$FileNamePath = "$PSSCriptRoot\RESULTS"
$FileName = "$FileNamePath\CREDITONEAPP.TST\*.txt","$FileNamePath\CREDITONEAPP.BIZ\*.txt","$FileNamePath\FNBM.CORP\*.txt","$FileNamePath\PHX.FNBM.CORP\*.txt"
if (Test-Path $FileName) {
  Remove-Item $FileName
}
$ResultsPath = $FileNamePath
#### Let's load a function
. .\Functions\Out-FileUtf8NoBom.ps1
. .\Functions\Get-SEPVersion.ps1
. .\Functions\Get-Server.ps1


#### FNBM.CORP
#Get-ADComputer -Server LASDC02.FNBM.CORP -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
#    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.txt"  -Append
Get-Server -FNBM | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.txt"  -Append
Get-Content "$ResultsPath\FNBM.CORP\FNBM.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\FNBM.CORP\FNBM.Dead.txt" -Append }}

#### PHX.FNBM.CORP
#Get-ADComputer -Server PHXDC03.PHX.FNBM.CORP -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
#    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.txt"  -Append
Get-Server -PHX | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.txt"  -Append
Get-Content "$ResultsPath\PHX.FNBM.CORP\PHX.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\PHX.FNBM.CORP\PHX.Dead.txt" -Append }}
  
#### C1B.BIZ
#Get-ADComputer -Server LASAUTH01.CREDITONEAPP.BIZ -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
#    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.txt"  -Append
Get-Server -BIZ | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.txt"  -Append
Get-Content "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Dead.txt" -Append }}

#### C1B.TST
#Get-ADComputer -Server LASAUTHTST01.CREDITONEAPP.TST -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
#    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.txt"  -Append
Get-Server -TST | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.txt"  -Append
Get-Content "$ResultsPath\CREDITONEAPP.TST\C1A.TST.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.TST\C1A.TST.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CREDITONEAPP.BIZ\C1A.TST.Dead.txt" -Append }}
  

Copy-Item "$ResultsPath\CREDITONEAPP.TST\C1A.TST.Alive.txt" -Destination "$ResultsPath\Alive\C1A.TST.Alive.txt"
Copy-Item "$ResultsPath\FNBM.CORP\FNBM.Alive.txt" -Destination "$ResultsPath\Alive\FNBM.Alive.txt"
Copy-Item "$ResultsPath\PHX.FNBM.CORP\PHX.Alive.txt" -Destination "$ResultsPath\Alive\PHX.Alive.txt"
Copy-Item "$ResultsPath\CREDITONEAPP.BIZ\C1A.BIZ.Alive.txt" -Destination "$ResultsPath\Alive\C1A.BIZ.Alive.txt"

Get-Content "$ResultsPath\Alive\FNBM.Alive.txt" |
    ForEach { Get-SEPVersion -ComputerName $_ | Select ComputerName, SEPProductVersion, SEPDefinitionDate }