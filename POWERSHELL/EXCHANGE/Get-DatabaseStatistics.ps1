Function Get-DatabaseStatistics {
    $Databases = Get-MailboxDatabase -Status
    foreach($Database in $Databases) {
        $DBSize = $Database.DatabaseSize
        $MBCount = (Get-Mailbox -Database $Database.Name).Count
        $MBAvg = $DBSize.ToBytes() / $MBCount            

        New-Object PSObject -Property @{
            Server = $Database.Server.Name
            DatabaseName = $Database.Name
            LastFullBackup = $Database.LastFullBackup
            MailboxCount = $MBCount
            "DatabaseSize (GB)" = "{0:n2}" -f ($DBSize.ToBytes() / 1GB)
            "AverageMailboxSize (MB)" = "{0:n2}" -f ($MBAvg / 1MB)
            "WhiteSpace (MB)" = "{0:n2}" -f ($Database.AvailableNewMailboxSpace.ToBytes() / 1MB)
        }
    }
}