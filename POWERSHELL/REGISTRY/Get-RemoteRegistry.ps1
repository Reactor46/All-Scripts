<#Get-ADComputer -SearchBase "OU=Test_Servers,OU=Servers,OU=Las_Vegas,DC=fnbm,DC=corp" -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} |
Select -ExpandProperty -Name | Out-File C:\LazyWinAdmin\Registry\Test-Servers.txt

GC -Path C:\LazyWinAdmin\Registry\Test-Servers.txt | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File C:\LazyWinAdmin\Registry\Alive.log -append
  } else { 
  write-output "$_ is Dead!!!" | Out-File C:\LazyWinAdmin\Registry\Dead.log -append}}#>

Get-Service -ComputerName LASTSTINTRA04 -Name RemoteRegistry
Get-Service -ComputerName LASDEVTOOLS02 -Name RemoteRegistry
Get-Service -ComputerName LASDEVTOOLS03 -Name RemoteRegistry
Get-Service -ComputerName LASCHATTST01 -Name RemoteRegistry
Get-Service -ComputerName PSTEST02 -Name RemoteRegistry
Get-Service -ComputerName LASMQ04 -Name RemoteRegistry
Get-Service -ComputerName LASTEST05WEB -Name RemoteRegistry
Get-Service -ComputerName LASTEST05MT -Name RemoteRegistry
Get-Service -ComputerName LASFPTST01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST05GUI -Name RemoteRegistry
Get-Service -ComputerName LASBIXGTST01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST01GUI -Name RemoteRegistry
Get-Service -ComputerName LASQA01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST01WEB -Name RemoteRegistry
Get-Service -ComputerName LASDEVTOOLS01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST01MT -Name RemoteRegistry
Get-Service -ComputerName LASNCACHE02TST -Name RemoteRegistry
Get-Service -ComputerName DCTEST01WEB -Name RemoteRegistry
Get-Service -ComputerName LASNCACHE01TST -Name RemoteRegistry
Get-Service -ComputerName LASJIRATST01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST02GUI -Name RemoteRegistry
Get-Service -ComputerName LASPROCOFACTST1 -Name RemoteRegistry
Get-Service -ComputerName DCTEST01GUI -Name RemoteRegistry
Get-Service -ComputerName LASBPTST01 -Name RemoteRegistry
Get-Service -ComputerName LASSFSUPG01 -Name RemoteRegistry
Get-Service -ComputerName LASETLTST01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST02MT -Name RemoteRegistry
Get-Service -ComputerName LASSQLTST03 -Name RemoteRegistry
Get-Service -ComputerName LASSQL2016TST -Name RemoteRegistry
Get-Service -ComputerName LASMEDIATST01 -Name RemoteRegistry
Get-Service -ComputerName LASTELCOTST01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST02WEB -Name RemoteRegistry
Get-Service -ComputerName LASTABTST01 -Name RemoteRegistry
Get-Service -ComputerName LASTABDEV01 -Name RemoteRegistry
Get-Service -ComputerName LASCODETST01 -Name RemoteRegistry
Get-Service -ComputerName LASSTG01GUI -Name RemoteRegistry
Get-Service -ComputerName LASSTG01MT -Name RemoteRegistry
Get-Service -ComputerName LASSTG01WEB -Name RemoteRegistry
Get-Service -ComputerName LASMCETST02 -Name RemoteRegistry
Get-Service -ComputerName LASAIR01 -Name RemoteRegistry
Get-Service -ComputerName LASAIRCC01 -Name RemoteRegistry
Get-Service -ComputerName LASINFRA04 -Name RemoteRegistry
Get-Service -ComputerName LASCARBON01 -Name RemoteRegistry
Get-Service -ComputerName LASSCUP01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST06GUI -Name RemoteRegistry
Get-Service -ComputerName LASTEST06MT -Name RemoteRegistry
Get-Service -ComputerName LASTEST06WEB -Name RemoteRegistry
Get-Service -ComputerName LASPDEV01WEB -Name RemoteRegistry
Get-Service -ComputerName LASSTGDEV01 -Name RemoteRegistry
Get-Service -ComputerName LASPDEV01MT -Name RemoteRegistry
Get-Service -ComputerName LASPDEV01GUI -Name RemoteRegistry
Get-Service -ComputerName LASSQLSPTST01 -Name RemoteRegistry
Get-Service -ComputerName LASSP01TST -Name RemoteRegistry
Get-Service -ComputerName LASSP02TST -Name RemoteRegistry
Get-Service -ComputerName LASSPSQL01TST -Name RemoteRegistry
Get-Service -ComputerName LASSP03TST -Name RemoteRegistry
Get-Service -ComputerName LASSPTST01 -Name RemoteRegistry
Get-Service -ComputerName LASDSMAPPTST01 -Name RemoteRegistry
Get-Service -ComputerName LASTEST03GUI -Name RemoteRegistry
Get-Service -ComputerName LASTEST03MT -Name RemoteRegistry
Get-Service -ComputerName LASTEST03WEB -Name RemoteRegistry
Get-Service -ComputerName LASTEST04GUI -Name RemoteRegistry
Get-Service -ComputerName LASTEST04MT -Name RemoteRegistry
Get-Service -ComputerName LASDCTEST01WEB -Name RemoteRegistry
Get-Service -ComputerName LASTEST04WEB -Name RemoteRegistry
Get-Service -ComputerName LASHF01GUI -Name RemoteRegistry
Get-Service -ComputerName LASDCTEST01GUI -Name RemoteRegistry
Get-Service -ComputerName LASHF01MT -Name RemoteRegistry
Get-Service -ComputerName LASHF01WEB -Name RemoteRegistry
Get-Service -ComputerName LASRL01GUI -Name RemoteRegistry
Get-Service -ComputerName LASSECSRV01 -Name RemoteRegistry
Get-Service -ComputerName IQORCOLLECT02 -Name RemoteRegistry
Get-Service -ComputerName LASTRITON01 -Name RemoteRegistry
Get-Service -ComputerName LASRL01MT -Name RemoteRegistry
Get-Service -ComputerName LASDCTEST01MT -Name RemoteRegistry
Get-Service -ComputerName LASRL01WEB -Name RemoteRegistry
Get-Service -ComputerName LASCADDTST01 -Name RemoteRegistry
Get-Service -ComputerName LASETLTST04 -Name RemoteRegistry
Get-Service -ComputerName LASMCETST01 -Name RemoteRegistry
Get-Service -ComputerName LASETLTST03 -Name RemoteRegistry