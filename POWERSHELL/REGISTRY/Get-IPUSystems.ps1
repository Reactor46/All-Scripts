<#### Let's delete the existing results.
$FileNamePath = "C:\LazyWinAdmin\Registry\RESULTS"
$FileName = "$FileNamePath\Desktops.txt","$FileNamePath\Alive.txt","$FileNamePath\Dead.txt"
if (Test-Path $FileName) {
  Remove-Item $FileName
}
$ResultsPath = $FileNamePath


#### FNBM.CORP
Get-ADComputer -Server LASDC02.FNBM.CORP -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom "$ResultsPath\Desktops.txt"  -Append

Get-Content "$ResultsPath\Desktops.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\Dead.txt" -Append }}
  #>
Get-Content "$ResultsPath\Alive.txt" |
    ForEach {Get-Service -Name  RemoteRegistry | Set-Service -StartupType Automatic -Status Running}

Get-Content "$ResultsPath\Alive.txt" |
    ForEach { If (Get-RegKeys -Hive HKLM -Key "SYSTEM\Setup\Upgrade\DownlevelBuildNumber" -ErrorAction SilentlyContinue) {write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\IPU.txt" -Append
    } else {
    write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\Clean.txt" -Append }}

Get-Content "$ResultsPath\Alive.txt" |
    ForEach {Get-Service -Name  RemoteRegistry | Set-Service -StartupType Disabled -Status Stopped}
