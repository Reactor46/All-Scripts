function CheckTime {

$servers = "LASDC01","LASDC02","LASDC05","PHXDC03","PHXDC04"

$tstats=@()

$timer = [diagnostics.stopwatch]::startnew()

foreach ($server in $servers){
$wdt = (gwmi -ComputerName $server -Query "select LocalDateTime from win32_operatingsystem").LocalDateTime
$dt = ([wmi]'').ConvertToDateTime($wdt) - $timer.elapsed

$tstat = "" |select Server,Timestamp
$tstat.server = $server
$tstat.timestamp = $dt
$tstats += $tstat
}

$enddate = (Get-Date).tostring("MM-dd-yyyy")
$filenameTXT = 'C:\LazyWinAdmin\WinSysChecklist\Logs\Time ' + $enddate + '.txt'
$filenameHTML = 'C:\LazyWinAdmin\WinSysChecklist\Logs\Time ' + $enddate + '.html'
#Then pick your poison
$tstats | Out-File $filenameTXT
$tstats | convertto-html | Out-File $filenameHTML
#End function CheckTime
}