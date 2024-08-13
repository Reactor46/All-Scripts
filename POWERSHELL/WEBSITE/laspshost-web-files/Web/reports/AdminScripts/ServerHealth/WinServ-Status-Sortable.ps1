
## Function to get the up time from the server
Function Get-UpTime
{
    param([string] $LastBootTime)
    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
    "$($Uptime.Days) days $($Uptime.Hours)h $($Uptime.Minutes)m"
}

## Begining of the loop. Lower down the loop is broken if the refresh option is not configured.

    ## Change value of the following parameter as needed
    $OutputFile = "C:\Scripts\Repository\jbattista\Web\reports\ServerStatusReportSortable.htm"
    $ServerList = Get-Content "C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\ServerHealth\Configs\Servers.txt" -ErrorAction SilentlyContinue # "LASMT03","LASMT10","LASPROCESS01"
    $Result = @()
    $TimeToRun=Measure-Command {
    ## Look through the servers in the file provided
    ForEach ($ServerName in $ServerList)
    {
        $PingStatus = Test-Connection -ComputerName $ServerName -Count 1 -Quiet

        ## If server responds...
        If ($PingStatus)
        {
            $OperatingSystem = Get-WmiObject Win32_OperatingSystem -ComputerName $ServerName
            
            $CpuUsage = Get-CIMInstance Win32_Processor -Computername $ServerName | Measure-Object -Property LoadPercentage -Average | ForEach-Object {"{0:N0}" -f $_.Average}

            $Uptime = Get-Uptime($OperatingSystem.LastBootUpTime)
            
            $MemUsage = Get-CIMInstance Win32_OperatingSystem -ComputerName $ServerName |
                            ForEach-Object {“{0:N0}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) * 100)/ $_.TotalVisibleMemorySize)}
                                            
            $IPv4Address = Get-CIMInstance Win32_NetworkAdapterConfiguration -ComputerName $ServerName | Select-Object -Expand IPAddress | Select -First 1 | Where-Object { ([Net.IPAddress]$_).AddressFamily -eq "InterNetwork" }
            
            $Service = Get-Process -ComputerName $ServerName -Name DataLayerService -ErrorAction SilentlyContinue | Select-Object  @{Name = "MemoryUsage"; Expression = {“{0:F0}” -f (($_.WS))}} | Sort -Property WS
            
            $Requests = Get-Counter -ComputerName $ServerName -Counter "\asp.net\requests queued" -ErrorAction SilentlyContinue                                      
            
            $TotalCon = Get-Counter -ComputerName $ServerName -Counter "\web service(_total)\current connections" -ErrorAction SilentlyContinue
                                        
	   
         } # End If ($PingStatus) Line 139
	
        ## Put the results together
        $Result += New-Object PSObject -Property @{
	        ServerName = $ServerName
		    IPV4Address = $IPv4Address
		    Status = $PingStatus
            CpuUsage = $CpuUsage
            Uptime = $Uptime
            MemUsage = $MemUsage
            Service = $Service.MemoryUsage
            Requests = $Requests.CounterSamples.CookedValue
            TotalCon = $TotalCon.CounterSamples.CookedValue
            
	    } # End $Result Line 173

        ## Clear the variables after obtaining and storing the results so offline servers don't have duplicate info.
        Clear-Variable IPv4Address
        Clear-Variable Uptime
        Clear-Variable MemUsage
        Clear-Variable CpuUsage
                
    } #End ForEach($ServerName in $ServerList) Line 134
     


$tr = @"
<tr class="w3-dark-grey item w3-small">
"@

$HTML = @" 
                
                <html>
                
                <head>
                
                <title> Server Health </title>
                
                </head>
                
                <meta name="viewport" content="width=device-width, initial-scale=1">
		        <link rel="stylesheet" href="../css/w3.css">
		        <script src="../js/w3.js"></script>
                
                <body class="w3-black">
                <div class="w3-container">                
                
                <table class="w3-table-all" id="sort">
                <thead>
                <tr class="w3-teal">
				<th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(1)')" style="cursor:pointer"><b><font color=#e6e6e6>Server</font></b></th>
                <th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(2)')" style="cursor:pointer"><b><font color=#e6e6e6>IP</font></b></th>
                <th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(3)')" style="cursor:pointer"><b><font color=#e6e6e6>Status</font></b></th>
                <th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(4)')" style="cursor:pointer"><b><font color=#e6e6e6>DLS Memory Usage</font></b></th>
                <th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(5)')" style="cursor:pointer"><b><font color=#e6e6e6>CPU Usage</font></b></th>
                <th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(6)')" style="cursor:pointer"><b><font color=#e6e6e6>Memory Usage</font></b></th>
                <th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(7)')" style="cursor:pointer"><b><font color=#e6e6e6>Queuing</font></b></th>
                <th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(8)')" style="cursor:pointer"><b><font color=#e6e6e6>Total Connections</font></b></th>                
                <th onclick="w3.sortHTML('#sort', '.item', 'td:nth-child(9)')" style="cursor:pointer"><b><font color=#e6e6e6>Uptime</font></b></th>
				</tr>
                </thead>
                
"@

    ## If there is a result put the HTML file together.
    If ($Result -ne $null){
              
        ## Highlight the alerts if the alerts are triggered.
        ## Font #00e600 = Green (Good)
        ## Symbol &#10004 = Green Check (Good/Online)
        ## Font #ffff4d = Yellow (Warning)
        ## Symbol &#9888 = Yellow Warning Marker (Warning) 
        ## Font #FF4D4D = Red (Alert)
        ## Symbol &#10008 = Red X (Alert/Offline)
        ForEach($Entry in $Result){  
         $HTML += $tr
            #Server Name
            If ($Entry.Status -eq $True)
            {
                $HTML += "<td><font color=#00e600>$($Entry.ServerName)</font></td>"
            }
            Else
            {
                $HTML += "<td><font color=#FF4D4D>&#10008 $($Entry.ServerName)</font></td>"
            }
            
            #IPV4 Address
            If ($Entry.Status -eq $True)
            {
                $HTML += "<td><font color=#00e600>$($Entry.IPV4Address)</font></td>"
            }
            Else
            {
                $HTML += "<td><font color=#FF4D4D>&#10008 $($Entry.IPV4Address) Offline</font></td>"
            }

            #Server Status (Online/Offline)
            If ($Entry.Status -eq $True)
            {
                $HTML += "<td><font color=#00e600>&#10004 Online</font></td>"
            }
            Else
            {
                $HTML += "<td><font color=#FF4D4D>&#10008 Offline</font></td>"
            }

            ## check DataLayerService Memory Usage
            If ($Entry.Service -ne $null)
            {
            
            If($Entry.Service/1MB -ge 1800)
                {
                    $HTML += "<td><font color=#FF4D4D>&#128561 $([math]::Round($Entry.Service/1MB)) MB</font></td>" ## Symbol &#128561 = face screaming in fear
                }       
            
            ElseIf(($Entry.Service/1MB -ge 1300) -and ($Entry.Service/1MB -lt 1799))
                {
                $HTML += "<td><font color=#ffff4d>&#9888 $([math]::Round($Entry.Service/1MB)) MB</font></td>" ## Symbol &#9888 = Yellow Warning Marker (Warning) 
                }
            
            ElseIf($Entry.Service/1MB -le 1299)
                {
                $HTML +="<td><font color=#00e600>&#10004 $([math]::Round($Entry.Service/1MB)) MB</font></td>" ## Symbol &#10004 = Green Check (Good/Online)
                }
            Else
                {
                $HTML += "<td><font color=#FF4D4D>&#10008 Check DataLayerService</font></td>" ## Symbol &#10008 = Red X (Alert/Offline)
                }
            }
            Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Check DataLayerService</font></td>"
                }
            #CPU Usage
            If ($Entry.CpuUsage -ne $null)
            {
            If ($Entry.CpuUsage -ge 90)
                {
                    $HTML += "<td><font color=#FF4D4D>&#128561 $($Entry.CpuUsage) %</font></td>" ## Symbol &#128561 = face screaming in fear
                }
            ElseIf(($Entry.CpuUsage -ge 70) -and ($Entry.CpuUsage -le 89))
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.CpuUsage) %</font></td>" ## Symbol &#9888 = Yellow Warning Marker
                }
            ElseIf($Entry.CpuUsage -le 69)
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.CpuUsage) %</font></td>" ## Symbol &#10004 = Green Check (Good/Online)
                }            
            Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Offline?</font></td>"
                }
             }
             Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Offline?</font></td>"
                }

            # Memory Usage
            If ($Entry.MemUsage -ne $null)
            {
            If($Entry.MemUsage -ge 90)
                {
                    $HTML += "<td><font color=#FF4D4D>&#128561 $($Entry.MemUsage) %</font></td>" ## Symbol &#128561 = face screaming in fear
                }
            ElseIf(($Entry.MemUsage -ge 70) -and ($Entry.MemUsage -le 89))
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.MemUsage) %</font></td>" ## Symbol &#9888 = Yellow Warning Marker
                }
            ElseIf($Entry.MemUsage -le 69)
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.MemUsage) %</font></td>" ## Symbol &#10004 = Green Check (Good/Online)
                }
            Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Offline?</font></td>"
                }
            }
            Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Offline?</font></td>"
                }
                                  
            # check ASP.NET Queued Requests
            If ($Entry.Requests -ne $null)
            {
            If ($Entry.Requests -ge 251)
                {
                    $HTML += "<td><font color=#FF4D4D>&#128561 $($Entry.Requests)</font></td>" ## Symbol &#128561 = face screaming in fear
                }
            ElseIf(($Entry.Requests -ge 100) -and ($Entry.Requests -le 250))
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.Requests)</font></td>" ## Symbol &#9888 = Yellow Warning Marker
                }
            ElseIf($Entry.Requests -le 99)
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.Requests)</font></td>" ## Symbol &#10004 = Green Check (Good/Online)
                }
            Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                }
            }
            Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                }
            
            # check Total Connections
            If ($Entry.TotalCon -ne $null)
            {
            If ($Entry.TotalCon -ge 500)
                {
                    $HTML += "<td><font color=#FF4D4D>&#128561 $($Entry.TotalCon)</font></td>" ## Symbol &#128561 = face screaming in fear
                }
            Elseif(($Entry.TotalCon -ge 300) -and ($Entry.TotalCon -le 499))
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.TotalCon)</font></td>" ## Symbol &#9888 = Yellow Warning Marker
                }
            ElseIf($Entry.TotalCon -le 299)
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.TotalCon)</font></td>" ## Symbol &#10004 = Green Check (Good/Online)
                }
            Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                }
            }
            Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                }

            # Uptime Status
            If ($Entry.Status -eq $True)
            {
                $HTML += "<td><font color=#00e600>$($Entry.Uptime)</font></td></tr>"
            }
            Else
            {
            $HTML += "<td><font color=#FF4D4D>&#10008 Offline</font></td></tr>"
            }

                
        
         } # End ForEach($Entry in $Result) Line 281
    } # End If ($Result -ne $null) Line 272
    }
    
        ## Report the date and time the script ran.
        $HTML += "</table><p><font color=#e6e6e6>Status refreshed on: $(Get-Date -Format G) Script Time " + $TimeToRun.TotalSeconds + " Seconds</font></p></body></html>"
                
     
       
        # End $TimeToRun=Measure-Command Line 132
        
        ## Output the HTML file
     

$HTML | Out-File $OutputFile

## If the refresh time option is not configured, stop the loop.
#Until ($RefreshTime -eq $null)

## End