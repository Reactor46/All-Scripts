$DomainGroup = "GS-IT-NOC-Advisor"
$LocalGroup  = "Administrators"
$Computer    = (Get-Content C:\LazyWinAdmin\Add-ToAdmin.txt)
$Domain      = $env:USERDOMAIN

([ADSI]"WinNT://$Computer/$LocalGroup,group").psbase.Invoke("Add",([ADSI]"WinNT://$Domain/$DomainGroup").path)