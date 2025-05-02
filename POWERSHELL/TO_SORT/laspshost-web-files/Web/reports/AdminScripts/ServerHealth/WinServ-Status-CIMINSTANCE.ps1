<#
    .EXAMPLE
    WinServ-Status.ps1 -List C:\foo\servers.txt -O C:\foo -DiskAlert 90 -CpuAlert 95 -MemAlert 85 -Refresh 120
    The script will execute using the list of servers and output a html report called WinServ-Status-Report.htm to C:\foo.
    The disk usage alert will highlight at 90% usage for any one drive, the CPU usage alert will highlight at 95% usage,
    and the memory usage alert will highlight at 85% usage. The script will re-run every 2 minutes.
#>

## Set up command line switches and what variables they map to
[CmdletBinding()]
Param(
    [parameter(Mandatory=$True)]
    [alias("List")]
    $ServerFile,
    [parameter(Mandatory=$True)]
    [alias("O")]
    $OutputPath,
    [alias("CpuAlert")]
    $CpuAlertThreshold,
    [alias("MemAlert")]
    $MemAlertThreshold,
    [alias("Requests")]
    $RequestAlertThreshold,
    [alias("Service")]
    $ServiceMemThreshold,
    [alias("TotalCon")]
    $TotalWebThreshold,
    [alias("TCPConMax")]
    $TCPConThreshold)


## Function to get the up time from the server


Function Get-UpTime
{
    param([string] $LastBootTime)
    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
    "$($Uptime.Days) days $($Uptime.Hours)h $($Uptime.Minutes)m"
}


    
    ## Change value of the following parameter as needed
    $OutputFile = "$OutputPath\ServerStatusReport.htm"
    $ServerList = Get-Content $ServerFile
    $Result = @()
    
    ## Look through the servers in the file provided
    $TimeToRun=Measure-Command {
    ForEach ($ServerName in $ServerList)
    {
        $PingStatus = Test-Connection -ComputerName $ServerName -Count 1 -Quiet

        ## If server responds, get uptime and disk info
        If ($PingStatus)
        {
            $OperatingSystem = Get-CimInstance Win32_OperatingSystem -ComputerName $ServerName
            $CpuAlert = $false
            $CpuUsage = Get-CimInstance Win32_Processor -Computername $ServerName |
                            Measure-Object -Property LoadPercentage -Average |
                                ForEach-Object {$_.Average; If($_.Average -ge $CpuAlertThreshold){$CpuAlert = $True}; "%"}
            $Uptime = Get-Uptime($OperatingSystem.LastBootUpTime)
            $MemAlert = $false
            $MemUsage = Get-CimInstance Win32_OperatingSystem -ComputerName $ServerName |
                            ForEach-Object {"{0:N0}" -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) * 100)/ $_.TotalVisibleMemorySize); If((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) * 100)/ $_.TotalVisibleMemorySize -ge $MemAlertThreshold){$MemAlert = $True}; "%"}
            $DiskAlert = $false
            $DiskUsage = Get-CimInstance Win32_LogicalDisk -ComputerName $ServerName |
                            Where-Object {$_.DriveType -eq 3} | ForEach-Object {$_.DeviceID, [Math]::Round((($_.Size - $_.FreeSpace) * 100)/ $_.Size); If([Math]::Round((($_.Size - $_.FreeSpace) * 100)/ $_.Size) -ge $DiskAlertThreshold){$DiskAlert = $True}; "%"}
            $IPv4Address = Get-CimInstance Win32_NetworkAdapterConfiguration -ComputerName $ServerName |
                                Select-Object -Expand IPAddress | Select-Object -First 1 |
                                    Where-Object { ([Net.IPAddress]$_).AddressFamily -eq "InterNetwork" }
            $ServiceAlert = $false
            $Service = Get-Process -ComputerName $ServerName -Name DataLayerService -ErrorAction SilentlyContinue |
                        ForEach-Object {"{0:N0}" -f ($_.WS/1MB); If([Math]::Round($_.WS/1MB) -ge $ServiceMemThreshold){$ServiceAlert = $True}; "MB"}
            $RequestsAlert = $false
            $Requests = Get-Counter -ComputerName $ServerName -Counter "\asp.net\requests queued" -ErrorAction SilentlyContinue |
                            ForEach-Object {($_.CounterSamples.CookedValue); If(($_.CounterSamples.CookedValue) -ge $RequestAlertThreshold){$RequestsAlert = $True}}
            $TotalConAlert = $false
            $TotalCon = Get-Counter -ComputerName $ServerName -Counter "\web service(_total)\current connections" -ErrorAction SilentlyContinue |
                            ForEach-Object {($_.CounterSamples.CookedValue); If(($_.CounterSamples.CookedValue) -ge $TotalWebThreshold){$TotalConAlert = $True}}
            $TCPConAlert = $false
            $TCPCon =  Get-Counter -ComputerName $ServerName -Counter "\TCPv4\Connections Established" -ErrorAction SilentlyContinue |
                           ForEach-Object {($_.CounterSamples.CookedValue); If(($_.CounterSamples.CookedValue) -ge $TCPConThreshold){$TCPConAlert = $True}}
         } # End If ($PingStatus)
	
        ## Put the results together
        $Result += New-Object PSObject -Property @{
	        ServerName = $ServerName
		    IPV4Address = $IPv4Address
		    Status = $PingStatus
            CpuUsage = $CpuUsage
            CpuAlert = $CpuAlert
		    Uptime = $Uptime
            MemUsage = $MemUsage
            MemAlert = $MemAlert
            Service = $Service
            ServiceAlert = $ServiceAlert
            Requests = $Requests
            RequestsAlert = $RequestsAlert
            TotalCon = $TotalCon
            TotalConAlert = $TotalConAlert
            TCPCon = $TCPCon
            TCPConAlert = $TCP
	    } # End $Result

        ## Clear the variables after obtaining and storing the results so offline servers don't have duplicate info.
        Clear-Variable IPv4Address
        Clear-Variable Uptime
        Clear-Variable MemUsage
        Clear-Variable CpuUsage
                
    } #End ForEach($ServerName in $ServerList)
     

$timer = @"
<script>

function refreshpage(interval, countdownel, totalel){
	var countdownel = document.getElementById(countdownel)
	var totalel = document.getElementById(totalel)
	var timeleft = interval+1
	var countdowntimer

	function countdown(){
		timeleft--
		countdownel.innerHTML = timeleft
		if (timeleft == 0){
			clearTimeout(countdowntimer)
			window.location.reload()
		}
		countdowntimer = setTimeout(function(){
			countdown()
		}, 1000)
	}

	countdown()
}

window.onload = function(){
	refreshpage(300, "countdown") // refreshpage(duration_in_seconds, id_of_element_to_show_result)
}

</script>

<div>Next <a href="javascript:window.location.reload()">refresh</a> in <b id="countdown"></b> seconds</div>
"@

 $HTML = @"
                <style type="text/css">
                p {font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;font-size:14px}
                p {color:#ffffff;}
                #Header{font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;}
                #Header td, #Header th {font-size:15px;text-align:left;border:1px solid #1a1a1a;padding:2px 2px 2px 7px;color:#ffffff;}
	            #Header th {font-size:16px;text-align:center;padding-top:5px;padding-bottom:4px;background-color:#4B4B4B;color:#ffffff;}
	            #Header tr.alt td {color:#ffffff;background-color:#eaf2d3;}
                #Header tr:nth-child(even) {background-color:#4B4B4B;}
                #Header tr:nth-child(odd) {background-color:#4B4B4B;}
                body {background-color: #1a1a1a;}
	            </style>
                <head><meta http-equiv="refresh" content="30">
                <script src="../js/sorttable.js"></script>
                <title> Server Health </title>
                </head>
"@
  

$HTML += @"
                <html><body>
                <table class="sortable" border=1 cellpadding=0 cellspacing=0 id=header>
                <thead>
                <tr>
                <th ><b><font color=#e6e6e6>Server</font></b></th>
                <th ><b><font color=#e6e6e6>IP</font></b></th>
                <th ><b><font color=#e6e6e6>Status</font></b></th>
                <th ><b><font color=#e6e6e6>DLS Memory Usage</font></b></th>
                <th ><b><font color=#e6e6e6>CPU Usage</font></b></th>
                <th ><b><font color=#e6e6e6>Memory Usage</font></b></th>
                <th ><b><font color=#e6e6e6>Queuing</font></b></th>
                <th ><b><font color=#e6e6e6>TCP Est Connections</font></b></th>
                <th ><b><font color=#e6e6e6>Total Connections</font></b></th>                
                <th ><b><font color=#e6e6e6>Uptime</font></b></th>
                </thead>
            </tr>
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
                $HTML += "<td><font color=#FF4D4D>&#10008 Offline</font></td>"
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

            # check DataLayerService Memory Usage
            If ($Entry.Service -ne $null)
            {
             If ($Entry.ServiceAlert -eq $True)
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.Service)</font></td>"
                }

                Else
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.Service)</font></td>"
                }
                    }
                Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Check DataLayerService</font></td>"
                }

            #CPU Usage
            If ($Entry.CpuUsage -ne $null)
            {
                If ($Entry.CpuAlert -eq $True)
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.CpuUsage)</font></td>"
                }

                Else
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.CpuUsage)</font></td>"
                }
                    }
        
                Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Offline</font></td>"
                }

            # Memory Usage
            If ($Entry.MemUsage -ne $null)
            {
                If ($Entry.MemAlert -eq $True)
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.MemUsage)</font></td>"
                }

                Else
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.MemUsage)</font></td>"
                }
                    }

                Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Offline</font></td>"
                }
            
            # check ASP.NET Queued Requests
            If ($Entry.Requests -ne $null)
            {
                If ($Entry.RequestsAlert -eq $True)
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.Requests)</font></td>"
                }

                Else
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.Requests)</font></td>"
                }
                    }                
                Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                }

            # check TCP Connections Established
            If ($Entry.TCPCon -ne $null)
            {
                If ($Entry.TCPConAlert -eq $True)
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.TCPCon)</font></td>"
                }

                Else
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.TCPCon)</font></td>"
                }
                    }                
                Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                }

            # check Total Connections
            If ($Entry.TotalCon -ne $null)
            {
                If ($Entry.TotalConAlert -eq $True)
                {
                    $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.TotalCon)</font></td>"
                }

                Else
                {
                    $HTML += "<td><font color=#00e600>&#10004 $($Entry.TotalCon)</font></td>"
                }
                    }
                Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                }

            

             #Uptime Status
            If ($Entry.Status -eq $True)
                {
                    $HTML += "<td><font color=#00e600>$($Entry.Uptime)</font></td></tr>"
                }

                Else
                {
                    $HTML += "<td><font color=#FF4D4D>&#10008 Offline</font></td></tr>"
                }

                
        
         } # End ForEach($Entry in $Result) 
    } # End If ($Result -ne $null) 
    
        }
        ## Report the date and time the script ran.
        $HTML += "</table><p><font color=#e6e6e6>Status refreshed on: $(Get-Date -Format G) Script Time " + $TimeToRun.TotalSeconds + " Seconds</font></p></body></html>"
         
 
    ## Output the HTML file
    $HTML | Out-File $OutputFile