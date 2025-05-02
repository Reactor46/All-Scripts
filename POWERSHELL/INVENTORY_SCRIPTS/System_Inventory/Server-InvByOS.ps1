
#### Contoso.CORP
Get-ADComputer -Server LASDC02.Contoso.CORP -Filter {Operatingsystem -Like 'Windows Server 2016*' -and Enabled -eq 'true'} -Properties * |
        Select-Object Name,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion |
        Export-CSV ".\AllWindowsContoso.csv" -Append -NoTypeInformation -Encoding UTF8
    

#### PHX.Contoso.CORP
Get-ADComputer -Server PHXDC03.PHX.Contoso.CORP -Filter {Operatingsystem -Like 'Windows Server 2016*' -and Enabled -eq 'true'} -Properties * |
        Select-Object Name,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion |
        Export-CSV ".\AllWindowsPHX.csv" -Append -NoTypeInformation -Encoding UTF8
      
#### C1B.BIZ
  Get-ADComputer -Server LASAUTH01.CREDITONEAPP.BIZ -Filter {Operatingsystem -Like 'Windows Server 2016*' -and Enabled -eq 'true'} -Properties * |
        Select-Object Name,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion |
        Export-CSV ".\AllWindowsBIZ.csv" -Append -NoTypeInformation -Encoding UTF8
    
#### C1B.TST
  Get-ADComputer -Server LASAUTHTST01.CREDITONEAPP.TST -Filter {Operatingsystem -Like 'Windows Server 2016*' -and Enabled -eq 'true'} -Properties * |
        Select-Object Name,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion |
        Export-CSV ".\AllWindowsTST.csv" -Append -NoTypeInformation -Encoding UTF8
      


