$Computers = Get-ADComputer -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'}

Foreach ($Computer in $Computers){
if (Test-Connection $Computer.Name -Count 1 -Quiet){
$Type = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer.Name).Model
$Serial = (Get-WmiObject -Class Win32_BIOS -ComputerName $Computer.Name).SerialNumber
$Result = New-Object PSObject -Property @{
Name = $Computer.Name
Serial = $Serial
$ADType = $Type.Model
}
$Result
Set-ADComputer -Identity $Computer.Name -Replace @{serialNumber = "$Serial"} -WhatIf
#Set-ADComputer -Identity $Computer.Name -Replace @{type = $ADType} -WhatIf
}
}