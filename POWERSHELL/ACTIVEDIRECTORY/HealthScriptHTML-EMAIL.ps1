Set-Location -LiteralPath 'D:\HealthScript\ExchangeAnalyzer\'
$enddate = (Get-Date).tostring("MM-dd-yyyy")
$AnalyzerReport = 'D:\HealthScript\ExchangeAnalyzer\Reports\ExchangeAnalyzer.'+ $enddate +'.html'
$HealthReport = 'D:\HealthScript\ExchangeAnalyzer\Reports\ExchangeHealth.'+ $enddate +'.html'


####******************* set email parameters ****************** ######
$from="Exchange2013@vegas.com"
$to="john.battista@vegas.com"
$smtpserver="mailrelay.vegas.com"
####*******************######################****************** ######

<#
# Connect to Exchange
Function GoGo-PSExch {

    param(
        [Parameter( Mandatory=$false)]
        [string]$URL="vvcexcmb01"
    )
    
    #$Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

    $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos #-Credential #$Credentials

    Import-PSSession $ExOPSession -AllowClobber

}
## End Connect to Exchange



GoGo-PSExch
#>
D:\HealthScript\ExchangeAnalyzer\Test-ExchangeServerHealth.ps1 -ReportMode -ReportFile $HealthReport
D:\HealthScript\ExchangeAnalyzer\Run-ExchangeAnalyzer.ps1 -FileName $AnalyzerReport

$Report = (GC $HealthReport), (GC $AnalyzerReport) | Out-String

Send-MailMessage -From $from  -To $to -SMTPServer $smtpserver  -Subject "Exchange 2013 Health and Analyzer Report" -Body $Report
