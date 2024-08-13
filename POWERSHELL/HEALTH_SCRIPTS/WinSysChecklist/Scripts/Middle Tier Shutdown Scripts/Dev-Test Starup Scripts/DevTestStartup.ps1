$tstEnvMTServers = "LASTEST01MT", "LASTEST02MT", "LASTEST03MT", "LASTEST04MT", "LASTEST05MT", "LASTEST06MT", "LASTRN01MT", "LASRL01MT", "LASSTG01MT", "LASHF01MT"
$tstEnvGUIServers = "LASTEST01GUI", "LASTEST02GUI", "LASTEST03GUI", "LASTEST04GUI", "LASTEST05GUI", "LASTEST06GUI", "LASTRN01GUI", "LASRL01GUI", "LASSTG01GUI", "LASHF01GUI"
$tstEnvWEBServers = "LASTEST01WEB", "LASTEST02WEB", "LASTEST03WEB", "LASTEST04WEB", "LASTEST05WEB", "LASTEST06WEB", "LASTRN01WEB", "LASRL01WEB", "LASSTG01WEB", "LASHF01WEB"
$tstEnvMCEServers = "LASMCETST01", "LASMCETST02"
$tstEnvCHATServers = "LASCHATTST01", "LASCHATTST02"

$tstMTServices = "ContosoApplicationProcessingService","ContosoApplicationParsingService", "ContosoApplicationImportService","CreditPullService",
"ContosoIPFraudCheckService","FromPPSExchangeFileWatcherService","BoardingService","CentralizedCacheService","ContosoDataLayerService","CreditOneBatchLetterRequestService","ContosoLPSService",
"NetConnectTransactionsSavingService","ContosoCheckRequestService","CreditOne.LogParser.Service","CollectionsAgentTimeService","ContosoFinCenService","ContosoQueueProcessorService","ContosoDebitCardHolderFileWatcher",
"ContosoIdentityCheckService","ContosoApplicationProcessingService","ContosoApplicationParsingService","ContosoApplicationImportService","CreditPullService","ContosoIPFraudCheckService",
"FromPPSExchangeFileWatcherService","BoardingService","FdrOutGoingFileWatcher","ValidationTriggerWatcher","CreditOne.CustomerNotifications.EmailService",
"CreditOne.CustomerNotifications.SMSService", "W3SVC"

$tstGUIServices = "W3SVC"
$tstWEBServices = "W3SVC" 
$tstMCEServices = "CreditEngine" 
$tstCHATServices = "WhosOn","WhosOnGateway","WhosOnQuery","WhosOnReports"


function validateMT(){

foreach ($server in $tstEnvMTServers){

        foreach ($service in $tstMTServices){
            $result = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
            $status = $result.Status
            Write-Host "$server $service $status"
        }

}

}


function startupMT(){

foreach ($server in $tstEnvMTServers){

        foreach ($service in $tstMTServices){
           #Start-Service -inputobject $(Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue) 
             $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Write-Host "$service is NOT FOUND on $server"}
             elseif ($testsvc.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc.WaitForStatus('Running','00:00:30')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Write-Host "$service is STOPPED on $server"
                    }
                 else
                    {
                     Write-Host "$service is RUNNING on $server"
                    }
                }
             else {Write-Host "$service is RUNNING on $server"}



        }

}

}


function validateGUI(){

foreach ($server in $tstEnvGUIServers){

        foreach ($service in $tstGUIServices){
            $result = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
            $status = $result.Status
            Write-Host "$server $service $status"
        }

}

}


function startupGUI(){

foreach ($server in $tstEnvGUIServers){

        foreach ($service in $tstGUIServices){
           #Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
             $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Write-Host "$service is NOT FOUND on $server"}
             elseif ($testsvc.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc.WaitForStatus('Running','00:00:30')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Write-Host "$service is STOPPED on $server"
                    }
                 else
                    {
                     Write-Host "$service is RUNNING on $server"
                    }
                }
             else {Write-Host "$service is RUNNING on $server"}



        }

}

}


function validateWEB(){

foreach ($server in $tstEnvWEBServers){

        foreach ($service in $tstWEBServices){
            $result = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
            $status = $result.Status
            Write-Host "$server $service $status"
        }

}

}


function startupWEB(){

foreach ($server in $tstEnvWEBServers){

        foreach ($service in $tstWEBServices){
           #Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
             $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Write-Host "$service is NOT FOUND on $server"}
             elseif ($testsvc.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc.WaitForStatus('Running','00:00:30')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Write-Host "$service is STOPPED on $server"
                    }
                 else
                    {
                     Write-Host "$service is RUNNING on $server"
                    }
                }
             else {Write-Host "$service is RUNNING on $server"}


        }

}

}

function validateMCE(){

foreach ($server in $tstEnvMCEServers){

        foreach ($service in $tstMCEServices){
            $result = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
            $status = $result.Status
            Write-Host "$server $service $status"
        }

}

}


function startupMCE(){

foreach ($server in $tstEnvMCEServers){

        foreach ($service in $tstMCEServices){
           #Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
             $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Write-Host "$service is NOT FOUND on $server"}
             elseif ($testsvc.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc.WaitForStatus('Running','00:00:30')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Write-Host "$service is STOPPED on $server"
                    }
                 else
                    {
                     Write-Host "$service is RUNNING on $server"
                    }
                }
             else {Write-Host "$service is RUNNING on $server"}



        }

}

}

function validateCHAT(){

foreach ($server in $tstEnvCHATServers){

        foreach ($service in $tstCHATServices){
            $result = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
            $status = $result.Status
            Write-Host "$server $service $status"
        }

}

}


function startupCHAT(){

foreach ($server in $tstEnvCHATServers){

        foreach ($service in $tstCHATServices){
           #Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
             $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Write-Host "$service is NOT FOUND on $server"}
             elseif ($testsvc.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $server -name $service) -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc = Get-Service -computername $server -name $service -ErrorVariable err -ErrorAction SilentlyContinue
                 $testsvc.WaitForStatus('Running','00:00:30')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Write-Host "$service is STOPPED on $server"
                    }
                 else
                    {
                     Write-Host "$service is RUNNING on $server"
                    }
                }
             else {Write-Host "$service is RUNNING on $server"}





        }

}

}


Write-Host "Enter a Selection below to Verify or Start Dev-Test Services"
Write-Host "1) Validate Startup of Dev-Test Services"
Write-Host "2) Startup All Dev-Test Services" 
$Selection = Read-Host

switch ($Selection){
    1 { validateMT
        validateGUI
        validateWEB
        validateMCE
        validateCHAT 
        }
    2 { startupMT
        startupGUI
        startupWEB
        startupMCE
        startupCHAT
        }
    default {
        Write-Host "Invalid Selection, Exiting"
        }
}

