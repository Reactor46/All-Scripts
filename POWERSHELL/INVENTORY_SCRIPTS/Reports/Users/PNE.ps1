#Author: Prabhat Nigam
#Microsoft Architect and CTO @ Golden Five Consulting
#Date: 08/10/2017
#Description: Extract Password Never Expires enabled user list and email to the configured user.
#Disclaimer: Please use as your own risk
#Please update SMTPHost, From,To, ReportPath, Subjet, $Body#

$DC1 = "LASDC02.fnbm.corp"
$DC2 = "PHXDC03.phx.fnbm.corp"
$DC3 = "LASAUTH01.creditoneapp.biz"
$DC4 = "LASAUTHTST01.creditoneapp.tst"

$Reportpath1 = "c:\LazyWinAdmin\Reports\Users\PNE-FNBM-$((Get-Date).ToString('MM-dd-yyyy')).csv"
$Reportpath2 = "c:\LazyWinAdmin\Reports\Users\PNE-PHX-FNBM-$((Get-Date).ToString('MM-dd-yyyy')).csv"
$Reportpath3 = "c:\LazyWinAdmin\Reports\Users\PNE-CreditOneApp.Biz-$((Get-Date).ToString('MM-dd-yyyy')).csv"
$Reportpath4 = "c:\LazyWinAdmin\Reports\Users\PNE-CreditOneApp.Tst-$((Get-Date).ToString('MM-dd-yyyy')).csv"


Get-ADUser -Server $DC1 -Filter * -Properties * | Where-Object { $_.passwordNeverExpires -eq "true" } |
    Where-Object {$_.enabled -eq "true"} |
        Select-Object Name,DistinguishedName,PasswordNeverExpires,@{Name="Lastlogon"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
                Export-Csv $Reportpath1 -force -NoTypeInformation
                
Get-ADUser -Server $DC2 -Filter * -Properties * | Where-Object { $_.passwordNeverExpires -eq "true" } |
    Where-Object {$_.enabled -eq "true"} |
        Select-Object Name,DistinguishedName,PasswordNeverExpires,@{Name="Lastlogon"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
                Export-Csv $Reportpath2 -force -NoTypeInformation

Get-ADUser -Server $DC3 -Filter * -Properties * | Where-Object { $_.passwordNeverExpires -eq "true" } |
    Where-Object {$_.enabled -eq "true"} |
        Select-Object Name,DistinguishedName,PasswordNeverExpires,@{Name="Lastlogon"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
                Export-Csv $Reportpath3 -force -NoTypeInformation

Get-ADUser -Server $DC4 -Filter * -Properties * | Where-Object { $_.passwordNeverExpires -eq "true" } |
    Where-Object {$_.enabled -eq "true"} |
        Select-Object Name,DistinguishedName,PasswordNeverExpires,@{Name="Lastlogon"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |
                Export-Csv $Reportpath4 -force -NoTypeInformation

