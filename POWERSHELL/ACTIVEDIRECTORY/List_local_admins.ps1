Get-ADComputer -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-file 'C:\LazyWinAdmin\Servers\Servers.txt' -Append utf8

Get-Content 'C:\LazyWinAdmin\Servers\Servers.txt' | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File C:\LazyWinAdmin\Servers\Alive.txt -append utf8
  } else { 
  write-output "$_" | Out-File C:\LazyWinAdmin\Servers\Dead.txt -append utf8}}



$servers= get-content 'C:\LazyWinAdmin\Servers\Alive.txt'
$output = 'c:\LazyWinAdmin\Servers\FNBMCORP-local_admin_output.csv' 
$results = @()

foreach($server in $servers)
{
$admins = @()
$group =[ADSI]"WinNT://$server/Administrators" 
$members = @($group.psbase.Invoke("Members"))
$members | foreach {
 $obj = new-object psobject -Property @{
 Server = $Server
 Admin = $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
 }
 $admins += $obj
 } 
$results += $admins
}
$results| Export-csv $Output -NoTypeInformation
