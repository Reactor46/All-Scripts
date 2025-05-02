#### Massive Server List
#### Vegas.com AD Forest
#### Vegas.com
#### Let's delete the existing results.
Import-Module "$PSScriptRoot\CredentialManager\CredentialManager.psm1"
$FileNamePath = "$PSScriptRoot\Inventory\Servers"
$FileName = "$FileNamePath\USON\*.txt"
if (Test-Path $FileName -ErrorAction SilentlyContinue) {
  Remove-Item $FileName -ErrorAction SilentlyContinue
}

$ResultsPath = $FileNamePath

#### CORP
Get-ADComputer -Server USONVSVRDC03 -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name  | Out-FileUtf8NoBom "$ResultsPath\USON\USON.txt"  -Append

Get-Content "$ResultsPath\USON\USON.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\USON\USON.Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\USON\USON.Dead.txt" -Append }}
  

Copy-Item "$ResultsPath\USON\USON.Alive.txt" -Destination "$ResultsPath\USON\USON_Completed.txt" -Force


Remove-Module CredentialManager
