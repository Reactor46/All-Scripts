$Base = "C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\Inventory\TXT-Results"
$FileNames = "$Base\*.txt"

    If (Test-Path $FileNames){
        Remove-Item $FileNames}

Get-ADComputer -Server LASDC02.Contoso.corp -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom $Base\Contoso.TXT  -Append 
Get-ADComputer -Server PHXDC03.phx.Contoso.corp -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom $Base\PHX.TXT  -Append
Get-ADComputer -Server LASAUTHTST01.creditoneapp.tst -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom $Base\C1B-TST.TXT  -Append 
Get-ADComputer -Server LASAUTH01.creditoneapp.biz -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom $Base\C1B-BIZ.TXT  -Append 

Get-Content $Base\Contoso.TXT,$Base\PHX.TXT,$Base\C1B-TST.TXT,$Base\C1B-BIZ.TXT |
    Group | where {$_.count -eq 1} | % {$_.group[0]} | Set-Content $Base\All-Combined.txt




Get-Content $Base\All-Combined.txt | 

 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom $Base\Alive.txt -append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom $Base\Dead.txt -append}}

CD "C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\Inventory\"
& .\PSSydiScript.ps1