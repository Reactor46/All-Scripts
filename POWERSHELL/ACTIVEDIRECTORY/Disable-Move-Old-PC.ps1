$Computers = Get-Content -Path .\Old.txt

ForEach ($Computer in $Computers)

{ $ADComputer = $null

$ADComputer = Get-ADComputer $Computer -Properties Description

If ($ADComputer)

{ Add-Content .\computers.log -Value "Found $Computer, disabling and moved to Disabled Computers OU"

Set-ADComputer $ADComputer -Description "Computer Disabled on $(Get-Date)" -Enabled $false

Move-ADObject $ADcomputer -targetpath “OU=DISABLED COMPUTERS,OU=COMPUTERS,OU=USON,DC=USON,DC=LOCAL”

}

Else

{ Add-Content .\computers.log -Value "$Computer not in Active Directory"

}

} 