
Get-ADComputer -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'}  -ErrorAction SilentlyContinue -Properties * | Select -ExpandProperty Name |
Out-File "$PSScriptRoot\All-Server-List.txt" -Encoding ascii