$db = "Students_01"
$filter = "'(Database -eq ''CN=Students_01,CN=Databases,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=DMACC,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=dmacc,DC=edu'')'"
$exportPath = "c:\exports\" + $db + ".csv"

Get-Mailbox -ResultSize Unlimited -SortBy DisplayName -Filter $filter | Select-Object DisplayName, Alias | Export-Csv $exportPath -NoTypeInformation