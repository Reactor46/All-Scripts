#Get-Content 'C:\LazyWinAdmin\AV Check\WindowsEmbedded.txt' | 
# ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File 'C:\LazyWinAdmin\AV Check\Alive.log' -append
#  } else { 
#  write-output "$_ is Dead!!!" | Out-File 'C:\LazyWinAdmin\AV Check\Dead.log' -append}}
  
  <#Get-Content 'C:\LazyWinAdmin\AV Check\Windows7.txt' | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File 'C:\LazyWinAdmin\AV Check\Alive.log' -append
  } else { 
  write-output "$_ is Dead!!!" | Out-File 'C:\LazyWinAdmin\AV Check\Dead.log' -append}}
  
  Get-Content 'C:\LazyWinAdmin\AV Check\Windows10.txt' | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File 'C:\LazyWinAdmin\AV Check\Alive.log' -append
  } else { 
  write-output "$_ is Dead!!!" | Out-File 'C:\LazyWinAdmin\AV Check\Dead.log' -append}}

  Get-Content 'C:\LazyWinAdmin\AV Check\phx.txt' | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File 'C:\LazyWinAdmin\AV Check\Alive.log' -append
  } else { 
  write-output "$_ is Dead!!!" | Out-File 'C:\LazyWinAdmin\AV Check\Dead.log' -append}}#>
 
  $Computers = Get-Content 'C:\LazyWinAdmin\AV Check\Alive.log'
    ForEach($comp in $Computers){Get-AVStatus -Computername $comp | Select ComputerName,Name,DefinitionStatus,RealTimeProtectionStatus }#| Export-CSV -Path 'C:\LazyWinAdmin\AV Check\AV-Report.csv' -Append -NoTypeInformation}


 