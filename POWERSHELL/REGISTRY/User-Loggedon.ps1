#Get-ADComputer -SearchBase 'OU=Collections,OU=Las_Vegas,DC=fnbm,DC=corp' -Filter {Operatingsystem -Like 'Windows*' -and Enabled -eq 'true'}  -ErrorAction SilentlyContinue -Properties *  | Select -ExpandProperty Name | Out-File "C:\LazyWinAdmin\ReadRegistry\collections.txt"
#$computers = Get-Content -Path "C:\LazyWinAdmin\ReadRegistry\collections.txt"
#$InactiveFile = "C:\LazyWinAdmin\ReadRegistry\Inactive.txt"
#$ActiveFile = "C:\LazyWinAdmin\ReadRegistry\Active.txt"
Get-Content "C:\LazyWinAdmin\ReadRegistry\Collections.txt" |

ForEach {
if ( -not (Test-Connection -comp $_ -Quiet)){
Write-Host "$_ is down" -ForegroundColor Yellow
} Else {
Out-File "C:\LazyWinAdmin\ReadRegistry\Alive.txt" -Append
}}