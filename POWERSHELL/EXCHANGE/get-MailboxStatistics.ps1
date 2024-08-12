Get-MailboxStatistics "scmiles" | fl TotalitemSize

Get-MailboxStatistics -Database Ankeny_Staff_01 | Sort-Object -descending TotalitemSize  |ft DisplayName,ItemCount,TotalItemSize


Get-MailboxStatistics -Database Students_01 | Sort-Object -descending TotalitemSize | Select-Object DisplayName,ItemCount,TotalItemSize | export-csv size.csv


Get-MailboxStatistics -Database Students_01 | Sort-Object -descending TotalitemSize | Select-Object DisplayName,ItemCount,@{n='TotalItemSize(MB)';e={$_.Totalitemsize.value.toMB()}} | export-csv size.csv