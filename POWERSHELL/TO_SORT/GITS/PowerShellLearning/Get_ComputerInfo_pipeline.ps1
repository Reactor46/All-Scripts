#Using piplines (|), to select a specific item of computer info (TimeZone)
Get-ComputerInfo | Select-Object -Property TimeZone