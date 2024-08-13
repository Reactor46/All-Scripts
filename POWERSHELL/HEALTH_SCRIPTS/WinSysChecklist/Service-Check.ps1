Set-Location -LiteralPath 'C:\LazyWinAdmin\WinSysChecklist'
$enddate = (Get-Date).tostring("MM-dd-yyyy")
$filename = 'C:\LazyWinAdmin\WinSysChecklist\Logs\Services.'+ $enddate +'.html'
#$filename2 = 'C:\LazyWinAdmin\WinSysChecklist\Logs\Time.'+$enddate +'.html'
$filename3 = 'C:\LazyWinAdmin\WinSysChecklist\Logs\DomainHealth.Contoso'+ $enddate +'.html'
$filename3A = 'C:\LazyWinAdmin\WinSysChecklist\Logs\DomainHealth.PHX.Contoso'+ $enddate +'.html'
$filename3B = 'C:\LazyWinAdmin\WinSysChecklist\Logs\DomainHealth.C1B.TST'+ $enddate +'.html'
$filename3C = 'C:\LazyWinAdmin\WinSysChecklist\Logs\DomainHealth.C1B.BIZ'+ $enddate +'.html'
$filename4 = 'C:\LazyWinAdmin\WinSysCheckList\Logs\ExchangeHealth.'+ $enddate +'.html'
$configs = 'C:\LazyWinAdmin\WinSysChecklist\Configs'
$W3SVC = Get-Content -Path $Configs\W3SVC.txt
$Contososvc = Get-Content -Path $Configs\Contosodls.txt
$filter = "name = 'DataLayerService.exe'"
$OutPut = @{Expression={$_.CSName};Label="Server"}, @{Label="Memory Usage";Expression={[String]([int]($_.WS/1MB))+" MB"}}, @{Expression={$_.CreationDate};Label="Running Since"}
$ReportTitle1= "CreditOne Bank Daily Server and Service Check Report $enddate"
$ReportTitle2= "CreditOne Bank W3SVC Report"
$ReportTitle3= "CreditOne Bank Contoso DataLayerService Memory Usage"
$ReportTitle4= "CreditOne Bank Contoso Collections Agent Time Service"
$ReportTitle5= "CreditOne Bank Credit Engine Service"
$ReportTitle6= "CreditOne Bank Credit Pull Service"
$ReportTitle7= "CreditOne Bank WhosOn Chat Service"
$ReportTitle8= "CreditOne Bank MS SQL SERVER & SQL SERVER AGENT Services"
$ReportTitle9= "CreditOne Bank P360 Services"
$ReportTitle10= "CreditOne Bank Service1 Services"
$ReportTitle11= "CreditOne Bank RFax Services"
$ReportTitle12= "CreditOne Bank AccuRev & JIRA Services"
$ReportTitle13= "CreditOne Bank Symantec EndPoint Protection Management Services"
$ReportTitle14= "CreditOne Bank Print Server Services"
$ReportTitle15= "CreditOne Bank Scheduler Service"
$ReportTitle16= "CreditOne Bank AD Time Check"
$npt = "LASDC01","LASDC02","LASDC03","LASDC10","LASDC11","PHXDC03","PHXDC04","LASAUTHTST01.creditoneapp.tst","LASAUTH01.creditoneapp.biz"

####******************* set email parameters ****************** ######
$from="winsys_service_chk@creditone.com"
$to="john.battista@creditone.com"
$smtpserver="mailgateway.Contoso.corp"
####*******************######################****************** ######

function CheckTime {

$tstats=@()

$timer = [diagnostics.stopwatch]::startnew()

foreach ($server in $npt){
$wdt = (Get-WmiObject -ComputerName $server -Query "select LocalDateTime from win32_operatingsystem").LocalDateTime
$dt = ([wmi]'').ConvertToDateTime($wdt) - $timer.elapsed

$tstat = "" |select Server,Timestamp
$tstat.server = $server
$tstat.timestamp = $dt
$tstats += $tstat
}

$enddate = (Get-Date).tostring("MM-dd-yyyy")
$tstats

#End function CheckTime
}

 
$Style = @"
 <header>
</header>
<style>
BODY{font-family:Calibri;font-size:12pt;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;color:black;background-color:#0BC68D;text-align:center;}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;text-align:center;}
</style>
"@
$query = @"
<QueryList>

  <Query Id="0" Path="Application">

    <Select Path="Application">*[System[Provider[@Name='MSExchangeIS']

    and (Level=3) and (EventID=10024) or (EventID=10025)]]

    </Select>

  </Query>

</QueryList>

"@

$frag2 = Get-Service -ComputerName $W3SVC  -Name W3SVC -ErrorAction SilentlyContinue|
    Select-Object MachineName, Status, Displayname |
        Sort MachineName |
            ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle2 </h2>" |
                Out-String
$frag3 = Get-CimInstance Win32_Process -ComputerName $Contososvc -Filter $filter -Property * -ErrorAction SilentlyContinue |
    Select $OutPut |
        Sort Server |
            ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle3 </h2>" |
                Out-String
$frag4 = Get-Service -Computername LASSVC03 -Name CollectionsAgentTimeService -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle4 </h2>" |
            Out-String
$frag5 = Get-Service -Computername LASMCE01, LASMCE02 -Name CreditEngine -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle5 </h2>" |
            Out-String
$frag6 = Get-Service -Computername LASCAPSMT01, LASCAPSMT02, LASCAPSMT05, LASCAPSMT06 -Name CreditPullService -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle6 </h2>" |
            Out-String
$frag7 = Get-Service -Computername LASCHAT01, LASCHAT02, LASCHAT03, LASCHAT04 -Name WhosOn* -ErrorAction SilentlyContinue|
    Select-Object MachineName, Status, Displayname |
        Select MachineName, Status, DisplayName |
            Where Status -eq "Running" |
                ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle7 </h2>" |
                    Out-String
$frag9 = Get-Service -Computername LASPROCDB02 -Name MSSQLSERVER,SQLSERVERAGENT -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle8 </h2>" |
            Out-String
$frag10 = Get-Service -Computername LASPROCAPP01 -Name P360* -ErrorAction SilentlyContinue|
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle9 </h2>" |
            Out-String
$frag11 = Get-Service -Computername LASPROCAPP04 -Name Service1 -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle10 </h2>" |
            Out-String
$frag12 = Get-Service -Computername LASRFAX01 -Name RF* -ErrorAction SilentlyContinue|
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle11 </h2>" |
            Out-String
$frag13 = Get-Service -Computername LASCODE02 -Name AccuRev*,JIRA050414101333 -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle12 </h2>" |
            Out-String
$frag14 = Get-Service -Computername LASITS02 -DisplayName *Symantec* -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle13 </h2>" |
            Out-String
$frag15 = Get-Service -Computername LASPRINT02 -Name Spooler -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle14 </h2>" |
            Out-String
$frag16 = Get-Service -Computername LASINFRA02 -Name Schedule -ErrorAction SilentlyContinue |
    Select-Object MachineName, Status, Displayname |
        ConvertTo-HTML -As Table -Fragment -PreContent "<h2> $ReportTitle15 </h2>" |
            Out-String
$frag17 = CheckTime |
    ConvertTo-Html -As Table -Fragment -PreContent "<h2> $ReportTitle16 </h2>" |
        Out-String


ConvertTo-HTML -head $style -PostContent $frag2,$frag3,$frag4,$frag5,$frag6,$frag7,$frag9,$frag10,$frag11,$frag12,$frag13,$frag14,$frag15,$frag16,$frag17 -PreContent "<h1> $ReportTitle1 </h1>" | Out-File $filename

.\chklst_DCDiag-HTML.ps1 -DomainName "Contoso.corp" -HTMLFileName $filename3
.\chklst_DCDiag-HTML.ps1 -DomainName "phx.Contoso.corp" -HTMLFileName $filename3A
.\chklst_DCDiag-HTML.ps1 -DomainName "creditoneapp.tst" -HTMLFileName $filename3B
.\chklst_DCDiag-HTML.ps1 -DomainName "creditoneapp.biz" -HTMLFileName $filename3C

GoGo-PSExch

C:\LazyWinAdmin\WinSysChecklist\ExchangeAnalyzer\Run-ExchangeAnalyzer.ps1 -FileName $filename4

$Body = (GC $filename), (GC $filename3), (GC $filename3A), (GC $filename3B), (GC $filename3C), (GC $filename4) | Out-String

Send-MailMessage -To $to -From $from -Subject "Daily Report" -Body $Body -BodyAsHtml -SmtpServer $smtpserver

Get-PSSession | Remove-PSSession