Get-ADUser -Server MWTDC02 -Filter * -Properties PasswordNeverExpires,msDS-UserPasswordExpiryTimeComputed | where {$_.enabled -eq $true -and $_.PasswordNeverExpires -eq  $False -and $_.name -Like "*svc*"} |
        select Name,@{Name="ExpiryDate";Expression={([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")).DateTime}} |
            where {($_.ExpiryDate | get-date)  -gt (get-date) -and ($_.ExpiryDate | get-date) -lt (get-date).adddays(90) } |
FT Name, ExpiryDate -AutoSize