Import-Module ActiveDirectory
$date = Get-Date -Format "MM_dd_yyyy_HH_mm"
$name = "Disabled_Accounts_Report_$date"
Search-ADAccount -AccountDisabled | Select-Object Name >> c:\LazyWinAdmin\Reports\$name.csv
"The Report of disabled user has been placed on the root of C:\LazyWinAdmin\Reports." 