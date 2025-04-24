###################################
# Web Server Monitor
# By: Josh D
# Created: 4/12
# Edit: 4/27
###################################

# AD Group populated with list of servers to test connectivity 
$servergroup = (Get-ADGroupMember -identity 'ServerMonitor').name | sort name

# AD Group populated with list of servers to reset IIS/websites upon failed connectivity test 
# to any of the above ServerMonitor group members (not currently needed to be a group, only 1 member)
$iisreset = (Get-ADGroupMember -Identity 'IISReset').name | sort name

# Variables
$log = "c:\scripts\monitorlogtest.txt"
$today = get-date
$errorcount = 0
$txtecount = 0
$error.clear()
$serverlist = ($servergroup  -join ", ")
$checkcount = $null
$tokentest = 0
[string[]]$recipients = "<>"
[string[]]$txtrecips = "<>"
$body = "Server Monitor Fault Report</br>"
$body+="-------------------------------</br></br>"
$txt = "IIS error(s): "


######### SERVER PING TEST #########

# Tests Server connections, if failed, waits 5s and tests again, if fails a second time, resets 3 IIS sites and tests a third time
$body+= "[INFO]`t Server Ping Monitor is currently monitoring: $Serverlist</br>" 
        "[INFO]`t Server Ping Monitor is currently monitoring: $Serverlist" | out-file $log -append 
Foreach($server in $servergroup)
  {
    if(!(Test-Connection -Cn $server -Count 1 -ea 0))
      {
        $body+= "[ERROR]`t Failure pinging $server at $today, will retry in 5 seconds<br/>"
                "[ERROR]`t Failure pinging $server at $today, will retry in 5 seconds" | out-file $log -append    
        Start-Sleep -seconds 5                                                                
        if(!(Test-Connection -Computername $server -Count 1 -ea 0))
          {
          $body+= "[ERROR]`t Retry failed, Server Monitor still unable to ping $server :"
                  "[ERROR]`t Retry failed, Server Monitor still unable to ping $server :" | out-file $log -append
          $errorcount+= 1
          Foreach ($iisserver in $iisreset)       # still need to reorganize and replace with function
            {
              $body+= '<ul style="list-style-type: square;">'
              $body+= "<li>Restarting Websites on $iisserver.</li>"
                 "[INFO]`t Restarting Websites on $iisserver." | out-file $log -append
              $invokereturn = invoke-command -computername "$iisserver" -scriptblock `
                {
                  import-module WebAdministration
                  stop-website ""
                  stop-website ""                  
                  stop-website ""
                  start-website ""
                  start-website ""
                  start-website "" 
                }
            }
          if(!(Test-Connection -Computername $server -Count 1 -ea 0))
            {  
              $errorcount+= 1
              $txtecount += 1
              $body+= "<li>Website reset is unable to correct problem with $server. Manual intervention is required.</li>" 
                "[ERROR]`t Website reset is unable to correct problem with $server. Manual intervention is required." | out-file $log -append
              $body+= "</ul>"
              $txt+=  "No connection to $server. "
            }
          else
            {
              $errorcount+= 1                                      # Disable this line to suppress email when IIS Reset has occured and fixed problem
              $body+= "<li>Website reset has corrected problem with $server. No further action is required.</li>" 
                 "[SUCCESS]Website reset has corrected problem with $server. No further action is required." | out-file $log -append
              $body+= "</ul>"
            }
          }
        else
          {
            $body+= "[SUCCESS]`tPing retry succeeded to $server.</br>" 
                    "[SUCCESS]`tPing retry succeeded to $server." | out-file $log -append
          }
      } 
  } 


######### CORE/PING PAGE TEST #########

# Poll site that checks uptime on .Auth, .WebApi.Resource, and .WebApi.Databee
# If statuscode is anything but 200, restart all 3 sites
try
  {
    $pingsite = invoke-webrequest 
    $pingstatus = $pingsite.StatusCode
    if ($pingsite.StatusCode -ne 200)                                      
      {
        $errorcount+= 1
        $body+= "[ERROR]`t Core/Ping page status code is $pingstatus, restarting IIS sites.</br>"
                "[ERROR]`t Core/Ping page status code is $pingstatus, restarting IIS sites." | out-file $log -append 
        Foreach ($iisserver in $iisreset)
          {
            $invokereturn = invoke-command -computername "$iisserver" -scriptblock `
              {
                import-module WebAdministration
                stop-website ""
                stop-website ""
                stop-website ""         
                start-website ""
                start-website ""
                start-website ""               
              }
          }
      }
  }
catch
  {
    $Error.Clear()
    $errorcount+= 1
    $txtecount += 1
    $body+= "[ERROR]`t Unable to communicate with Core/Ping page: $Error</br>"
            "[ERROR]`t Unable to communicate with Core/Ping page: $Error" | out-file $log -append 
    $txt+=  "Can't reach Core/Ping page. "
  }
  
          
######### TOKEN REQUEST TEST #########

# Request token, if token grant fails for any reason, restart all 3 sites 
$reqbody = @{
    Client_id = ""
    Client_secret = ""
    Username = ""
    Password = ""
    grant_type = ""
  }
try 
  {
    $token = Invoke-RestMethod -Uri http://172.18.0.16:8080/oauth/token -ContentType 'application/json' -body $reqbody -Method POST 
  }
catch 
  {
    $error.Clear()
    $errorcount+= 1
    $body+= "[ERROR]`t Unable to receive token, restarting IIS. StatusCode: "+$_.Exception.Response.StatusCode.value__+" / StatusDescription: "+$_.Exception.Response.StatusDescription+"</br>"
            "[ERROR]`t Unable to receive token, restarting IIS. StatusCode: "+$_.Exception.Response.StatusCode.value__+" / StatusDescription: "+$_.Exception.Response.StatusDescription+"</br>" | out-file $log -append
    Foreach ($iisserver in $iisreset)
      {
        $invokereturn = invoke-command -computername "$iisserver" -scriptblock `
          {
            import-module WebAdministration
            stop-website ""
            stop-website ""
            stop-website ""         
            start-website ""
            start-website ""
            start-website ""               
          }
      }

    # try to get token again, after restarting IIS
    try 
      {
        $token = Invoke-RestMethod -Uri http://172.18.0.16:8080/oauth/token -ContentType 'application/json' -body $reqbody -Method POST 
      }
    catch 
      {
        $body+= "[ERROR]`t Still unable to receive token after IIS restart. StatusCode: "+$_.Exception.Response.StatusCode.value__+" / StatusDescription: "+$_.Exception.Response.StatusDescription+"</br>"
                "[ERROR]`t Still unable to receive token after IIS restart. StatusCode: "+$_.Exception.Response.StatusCode.value__+" / StatusDescription: "+$_.Exception.Response.StatusDescription+"</br>" | out-file $log -append
        $txt+= "Can't get token. "
        $txtecount+= 1
        $tokentest = 1
      }
  }


######### QUEUE/COUNT PAGE TEST #########

# Skip test if token grant test failed. If passed, check Count page with token
If ($tokentest -eq 0)
  {  
    $bearerAuthValue = "Bearer "+$token.access_token
    $pingheaders = @{ 
        Authorization = $bearerAuthValue
      }

        # If status code is anything but 200, restart WebAPI.Resource
    Try
      {
        $checkcount = Invoke-WebRequest -uri  -Headers $pingheaders 
        if ($checkcount.StatusCode -ne 200)
          {
            $errorcount+= 1
            $countstatus = $checkcount.StatusCode
            $body+= "[ERROR]`t Queue/Count page status code is $countstatus. </br>"
                    "[ERROR]`t Queue/Count page status code is $countstatus." | out-file $log -append 
            Foreach ($iisserver in $iisreset)
              {
                $invokereturn = invoke-command -computername "$iisserver" -scriptblock `
                  {
                    import-module WebAdministration
                    stop-website "TSVG.WebAPI.Resource"
                    start-website "TSVG.WebAPI.Resource"
                  }
              }
          }
      }
    Catch
      {
        $errorcount+= 1
        $txtecount += 1
        $body+= "[ERROR]`t Unable to check Queue/Count page: $Error</br>"  
                "[ERROR]`t Unable to check Queue/Count page: $Error" | out-file $log -append 
        $txt+=  "Can't reach Queue/Count page. "
      } 

    If ($checkcount -ne $null)      # Converting return into table
      {    
        $checktrim = $checkcount.content.trim([Char]"{",[Char]"}") -replace '"','' -replace ':','='
        $finalcount = $checktrim.substring(0, $checktrim.Indexof(','))
        $isGT150 = convertfrom-stringdata $finalcount

        # If queuedCount is greater than 150, restart TSVG.WinService (display message for testing)
        $isGT150.keys | ForEach-Object `
          {
            $countme = '{1}' -f $_, $isGT150[$_]
            $countmessage = 'The {0} is {1}' -f $_, $isGT150[$_]
          }
        If ($countmessage -ne $null)
          {
            $body+= "[INFO]`t $countmessage at $today.</br>"  
                    "[INFO]`t $countmessage at $today." | out-file $log -append
          }
        If ($countme -gt 150)    # Set to -eq 0 for email testing
          {
            $body+= "[ERROR]`t The queuedCount is greater than 150:</br>"  
                    "[ERROR]`t The queuedCount is greater than 150:" | out-file $log -append
            $body+= '<ul style="list-style-type: square;">'
            $body+= '<li>Restarting the TSVG.WindowsService application.</li>'
               '[INFO]`t Restarting the TSVG.WindowsService application.'   | out-file $log -append

            # Restart TSVG.WinService on SVFF-CF-DEV using preconfigured scheduled task
            $servicereturn = invoke-command -computername "svff-cf-dev" -scriptblock `
              {
                Try
                  {
                    $stopproc = get-process -name "TSVG.WindowsService" | stop-process -Force     
                  }
                Catch{}
                Try
                  { 
                    $startproc = schtasks /run /TN "Call TSVG.WindowsService"                     
                  }
                Catch{}
              }
            $body+= "<li>The TSVG.WindowsService application has been restarted, waiting 5 minutes and rechecking queuedCount.</li>"
               "[INFO]`t The TSVG.WindowsService application has been restarted, waiting 5 minutes and rechecking queuedCount." | out-file $log -append
            $body+= "</ul>"

            # Pause 5 minutes and recheck the queuedCount website now that the application has been restarted
            start-sleep -Seconds 300
            Try                                # Checking status code is 200 again, restarting WebAPI.Resource if not
              {
                $checkcount = Invoke-WebRequest -uri  -Headers $pingheaders 
                if ($checkcount.StatusCode -ne 200)
                  {
                    $errorcount+= 1
                    $countstatus = $checkcount.StatusCode
                    $body+= "[ERROR]`t Queue/Count page status code is $countstatus. </br>"
                            "[ERROR]`t Queue/Count page status code is $countstatus." | out-file $log -append 
                    Foreach ($iisserver in $iisreset)
                      {
                        $invokereturn = invoke-command -computername "$iisserver" -scriptblock `
                          {
                            import-module WebAdministration
                            stop-website "TSVG.WebAPI.Resource"
                            start-website "TSVG.WebAPI.Resource"
                          }
                      }
                  }
              }
            Catch                                   # If unable to communicate with Queue/Count page after WinService restart 
              {
                $error.clear()
                $errorcount+= 1
                $txtecount += 1
                $body+= "[ERROR]`t Unable to check Queue/Count page after TSVG.WinService restart: $Error</br>"  
                        "[ERROR]`t Unable to check Queue/Count page after TSVG.WinService restart: $Error" | out-file $log -append 
                $txt+=  "Can't reach Queue/Count page. "
              }  

            $checktrim = $checkcount.content.trim([Char]"{",[Char]"}") -replace '"','' -replace ':','='
            $finalcount = $checktrim.substring(0, $checktrim.Indexof(','))
            $isGT150 = convertfrom-stringdata $finalcount

            # If queuedCount is still greater than 150 after corrective action, log the error 
            $isGT150.keys | ForEach-Object `
              {
                $countme = '{1}' -f $_, $isGT150[$_]
                $countmessage = 'the {0} is {1}' -f $_, $isGT150[$_]
              } 
            If ($countme -gt 150)    # Set to -eq 0 for email testing
              {
                $errorcount+= 1
                $txtecount += 1
                $body+= "[ERROR]`t After restarting application, $countmessage. Manual intervention is required."  
                        "[ERROR]`t After restarting application, $countmessage. Manual intervention is required." | out-file $log -append
                $txt+=  "$countmessage. "         
              }
            else
              {
          $body+= "[SUCCESS]`t Website reset has corrected the problem, $countmessage. No further action is required."
                  "[SUCCESS]`t Website reset has corrected the problem, $countmessage. No further action is required." | out-file $log -append
              }
          }
      }
  }


######### EMAIL TEST RESULTS #########

# Sends email only if errors were encountered 
If ($errorcount -ge 1)   # Set to 0 for testing so it always sends email
  {
    If ($errorcount -eq 1)
      {
      $subject = "$errorcount error experienced by the Server Monitor"
      }
      Else
        {
          $subject = "$errorcount errors experienced by the Server Monitor"
        }
    Send-MailMessage -smtp svff-mail-03 -from ServerMonitor@tsvg.com -to $recipients -subject $subject -bodyashtml $body
  }
If ($txtecount -ge 1)   # Set to 0 for testing so it always sends txts
  {
    $subject = "$txtecount"
    Send-MailMessage -smtp svff-mail-03 -from ServerMonitor@tsvg.com -to $txtrecips -subject $subject -body $txt
  }