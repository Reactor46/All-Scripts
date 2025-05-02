$Services = "MSExchangeADTopology","MSExchangeIS","MSExchangeMailboxAssistants","MSExchangeMailSubmission","MSExchangeMonitoring","MSExchangeRepl","MSExchangeRPC","MSExchangeSA","MSExchangeSearch","MSExchangeServiceHost","MSExchangeThrottling","MSExchangeTransportLogSearch"

ForEach ($Service in $Services) {
Set-Service -ComputerName USONVSVRDAG01 -Name $Service -StartupType Automatic -Status Running
}