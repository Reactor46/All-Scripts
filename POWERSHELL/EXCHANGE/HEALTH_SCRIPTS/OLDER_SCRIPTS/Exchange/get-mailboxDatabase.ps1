Get-MailboxDatabase | Where-Object {$_.Server -eq "ANEXMBX01"}

Get-MailboxDatabase | where{$_.Name -like "Students*"} | ft Name,DeletedItemRetention


Get-MailboxDatabase -Status | select ServerName,Name,DatabaseSize | where{$_.Name -like "Students*"} |sort-object DatabaseSize -Descending