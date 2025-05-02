### CSS style
$css= "<style>"
$css= $css+ "BODY{ text-align: center; background-color:white;}"
$css= $css+ "TABLE{    font-family: 'Lucida Sans Unicode', 'Lucida Grande', Sans-Serif;font-size: 12px;margin: 10px;width: 100%;text-align: center;border-collapse: collapse;border-top: 7px solid #004466;border-bottom: 7px solid #004466;}"
$css= $css+ "TH{font-size: 13px;font-weight: normal;padding: 1px;background: #cceeff;border-right: 1px solid #004466;border-left: 1px solid #004466;color: #004466;}"
$css= $css+ "TD{padding: 1px;background: #e5f7ff;border-right: 1px solid #004466;border-left: 1px solid #004466;color: #669;hover:black;}"
$css= $css+  "TD:hover{ background-color:#004466;}"
$css= $css+ "</style>" 
 
#$StartDate = (get-date).adddays(-1)
 
#$body = Get-WinEvent -FilterHashtable @{logname="Application"; starttime=$StartDate} -ErrorAction SilentlyContinue
 


#Written by David Maugrion -- david.maugrion@atos.net -- ATOS
#Change the idnumber(s) by the one you are searching

$ErrorActionPreference = "SilentlyContinue"

#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

$lognames = Get-winevent -ListLog *|select -expand logname;$total_logs=$lognames.Count


# Change this value(s) to what you are looking for

$idnumbers=(9041,3049,1012,1024,1135,1177,1069)

#how many day(s) back are we looking at

$d=-30

$servers = get-exchangeserver| select -expand name
$total_servers=$servers.count

"Will look in $total_logs Journal Logs for EventID(s) $idnumbers in the last ($d) on the following $total_servers server(s) $servers"

foreach ($server in $servers)
{
Write-Host "$Server" -foregroundcolor red -backgroundcolor yellow
$i=0
foreach ($logname in $lognames)
{ 
Write-Progress -Activity "Working on Server $server" -status "Looking in Log: $logname" -percentComplete ($i++ / $total_logs*100)
foreach ($idnumber in $idnumbers)
{

$body = Get-WinEvent -cn $server -FilterHashtable @{logname=$logname; StartTime =(Get-Date).AddDays($d);id=$idnumber}
}
}
}

$body | ConvertTo-HTML -Head $css MachineName,ID,TimeCreated,Message > C:\LazyWinAdmin\Exchange\ExchangeEvent-Errors.html