#### Massive Server List
#### Vegas.com AD Forest
#### Vegas.com
#### Let's delete the existing results.
Import-Module "C:\LazyWinAdmin\VDC-Domains\Inventory\CredentialManager\CredentialManager.psm1"
$FileNamePath = "C:\LazyWinAdmin\VDC-Domains\Inventory\Servers"
$FileName = "$FileNamePath\RES\*.txt","$FileNamePath\SVC\*.txt","$FileNamePath\CORP\*.txt"
if (Test-Path $FileName -ErrorAction SilentlyContinue) {
  Remove-Item $FileName -ErrorAction SilentlyContinue
}

$ResultsPath = $FileNamePath

#### CORP
Get-ADComputer -Server CORP-DC01 -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name  | Out-FileUtf8NoBom "$ResultsPath\CORP\CORP.txt"  -Append

Get-Content "$ResultsPath\CORP\CORP.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\CORP.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\CORP.Dead.txt" -Append }}
  
#### SVC
  Get-ADComputer -Server SVC-DC01 -Credential (Get-StoredCredential SVC) -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name | Out-FileUtf8NoBom "$ResultsPath\SVC\SVC.txt"  -Append

Get-Content "$ResultsPath\SVC\SVC.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\SVC\SVC.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\SVC\SVC.Dead.txt" -Append }}

#### RES
  Get-ADComputer -Server RES-DC01 -Credential (Get-StoredCredential RES) -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name | Out-FileUtf8NoBom "$ResultsPath\RES\RES.txt"  -Append

Get-Content "$ResultsPath\RES\RES.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\RES\RES.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\RES\RES.Dead.txt" -Append }}


Copy-Item "$ResultsPath\RES\RES.Alive.txt" -Destination "C:\LazyWinAdmin\VDC-Domains\Configs\RES.txt" -Force
Copy-Item "$ResultsPath\CORP\CORP.Alive.txt" -Destination "C:\LazyWinAdmin\VDC-Domains\Configs\CORP.txt" -Force
Copy-Item "$ResultsPath\SVC\SVC.Alive.txt" -Destination "C:\LazyWinAdmin\VDC-Domains\Configs\SVC.txt" -Force


Remove-Module CredentialManager

