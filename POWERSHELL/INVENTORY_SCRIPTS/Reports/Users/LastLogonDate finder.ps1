Import-Module ActiveDirectory
Get-ADuser -filter * -Properties *|
    where {$_.enabled -eq $False -and $_.LastLogonDate -ge (Get-Date).AddDays(-120)} |
        Select Name, LastLogonDate |
            Export-Csv -Path .\120output.csv -NoTypeInformation