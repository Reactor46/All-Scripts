Get-ExchangeServer | Format-List Name, Edition, AdminDisplayVersion
 
Get-Command ExSetup.exe | ForEach-Object { $_.FileVersionInfo }