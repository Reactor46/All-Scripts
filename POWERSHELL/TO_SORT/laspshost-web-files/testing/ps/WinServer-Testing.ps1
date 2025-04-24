<#PSScriptInfo

.VERSION 1.3.C.0.0.7

.AUTHOR John Battista, Based on code by Mike Galvin, Dan Price & Bhavik Solanki.

.LASTEDIT 1/15/2018 8:04 AM

.COMPANYNAME

.COPYRIGHT (C) Mike Galvin. All rights reserved.

.TAGS Windows Server Status Report Monitor

.LICENSEURI

.ORIGINAL_PROJECTURI https://gal.vin/2017/07/28/windows-server-status/

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

#>

<#
    .SYNOPSIS
    Creates a status report of Windows Servers.

    .DESCRIPTION
    Creates a status report of Windows Servers.

    This script will:
    
    Generate a HTML status report from a configurable list of Windows servers.
    The report will highlight information is the alert threashold is exceeded.

    Please note: to send a log file using ssl and an SMTP password you must generate an encrypted
    password file. The password file is unique to both the user and machine.
    
    The command is as follows:

    $creds = Get-Credential
    $creds.Password | ConvertFrom-SecureString | Set-Content c:\foo\ps-script-pwd.txt
    
    .PARAMETER List
    The path to a text file with a list of server names to monitor.

    .PARAMETER O
    The path where the HTML report should be output to. The filename will be WinServ-Status-Report.htm.

    .PARAMETER DiskAlert
    The percentage of disk usage that should cause the disk space alert to be raised.

    .PARAMETER CpuAlert
    The percentage of CPU usage that should cause the CPU alert to be raised.

    .PARAMETER MemAlert
    The percentage of memory usage that should cause the memory alert to be raised.

    .PARAMETER Refresh
    The number of seconds that she script should wait before running again. If not configured the script will run once and then exit.

    .PARAMETER SendTo
    The e-mail address the log should be sent to.

    .PARAMETER From
    The from address the log should be sent from.

    .PARAMETER Smtp
    The DNS or IP address of the SMTP server.

    .PARAMETER User
    The user account to connect to the SMTP server.

    .PARAMETER Pwd
    The password for the user account.

    .PARAMETER UseSsl
    Connect to the SMTP server using SSL.

    .EXAMPLE
    WinServ-Status.ps1 -List C:\foo\servers.txt -O C:\foo -DiskAlert 90 -CpuAlert 95 -MemAlert 85 -Refresh 120
    The script will execute using the list of servers and output a html report called WinServ-Status-Report.htm to C:\foo.
    The disk usage alert will highlight at 90% usage for any one drive, the CPU usage alert will highlight at 95% usage,
    and the memory usage alert will highlight at 85% usage. The script will re-run every 2 minutes.
#>

## Set up command line switches and what variables they map to.
## The parameters can be hardcoded. Comment out from Here:
[CmdletBinding()]
Param(
    [parameter(Mandatory=$True)]
    [alias("L")]
    $ServerFile,
    [parameter(Mandatory=$True)]
    [alias("O")]
    $OutputPath,
    [alias("CPUW")]
    $CPUWarnThreshold,
    [alias("CPUA")]
    $CPUAlertThreshold,
    [alias("MW")]
    $MemWarnThreshold,
    [alias("MA")]
    $MemAlertThreshold,
    [alias("RW")]
    $RequestWarnThreshold,
    [alias("RA")]
    $RequestAlertThreshold,
    [alias("SMW")]
    $ServiceMemWarnThreshold,
    [alias("SMA")]
    $ServiceMemAlertThreshold,
    [alias("CW")]
    $ConnWarnThreshold,
    [alias("CA")]
    $ConnAlertThreshold)
## .... to here

<#  << remove this comment to hardcode
##########################
## HardCoded Parameters ##
##########################
 
$L = .\Configs\Test.txt
$O = C:\inetpub\wwwroot\testing\reports\
$CPUW = 80
$CPUA = 90
$MW = 80
$MA = 90
$RW = 30
$RA = 100
$SMW = 1650
$SMA = 2000
$CW = 225
$CA = 300
#> # << remove this comment to hardcode



## Function to get the up time from the server

Function Get-UpTime
{
    param([string] $LastBootTime)
    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
    "$($Uptime.Days) days $($Uptime.Hours)h $($Uptime.Minutes)m"
}

## Begining of the loop. Lower down the loop is broken if the refresh option is not configured.
Do
{
    ## Change value of the following parameter as needed
    $OutputFile = "$OutputPath\WinServTesting.htm"
    $ServerList = Get-Content $ServerFile
    $Result = @()
    $TimeToRun=Measure-Command {
    ## Look through the servers in the file provided
    ForEach ($ServerName in $ServerList)
    {
        $PingStatus = Test-Connection -ComputerName $ServerName -Count 1 -Quiet

        ## If server responds, get uptime and disk info
        If ($PingStatus)
        {
            
            $OperatingSystem = Get-WmiObject Win32_OperatingSystem -ComputerName $ServerName
            
            $IPv4Address = Get-CimInstance Win32_NetworkAdapterConfiguration -ComputerName $ServerName |
                                Select-Object -Expand IPAddress | Select-Object -First 1 |
                                    Where-Object { ([Net.IPAddress]$_).AddressFamily -eq "InterNetwork" }
            
            
            $CpuUsage = Get-CimInstance Win32_Processor -Computername $ServerName | Measure-Object -Property LoadPercentage -Average |
                            ForEach {If(($CPUUsage.Average -ge $CPUWarnThreshold) -and ($CPUUsage.Average -le $CPUAlertThreshold -1)){$CpuWarn = $True}; If($CPUUsage.Average -ge $CPUAlertThreshold){$CPUAlert = $True}"%"}
            
            $Uptime = Get-Uptime($OperatingSystem.LastBootUpTime)                                           
            
            
            $MemUsage = ([Math]::Round(((($OperatingSystem.TotalVisibleMemorySize - $OperatingSystem.FreePhysicalMemory) * 100)/ $OperatingSystem.TotalVisibleMemorySize))) |
                            ForEach-Object {($MemUsage); If ($MemUsage -ge $MemWarnThreshold){$MemWarn = $True}; If ($MemUsage -ge $MemAlertThreshold){$MemAlert = $True}; If ($MemAlert = $True){$MemWarn = $False} "%"}
                                                                  
            iceAlert = $false
            $ServiceWarn = $false
            $Service = Get-Process -ComputerName $ServerName -Name DataLayerService -ErrorAction SilentlyContinue |
                           ForEach-Object {"{0:N0}" -f ($_.WS/1MB); If ([Math]::Round(($_.WS)/1MB) -ge $ServiceMemWarnThreshold){$ServiceWarn = $True};`
                               If ([Math]::Round(($_.WS)/1MB) -ge $ServiceMemAlertThreshold){$ServiceAlert = $True}; If ($ServiceAlert = $True){$ServiceWarn = $False} "MB"}
                        
                   
            $RequestsAlert = $false
            $RequestsWarn = $false
            $Requests = Get-Counter -ComputerName $ServerName -Counter "\asp.net\requests queued" -ErrorAction SilentlyContinue |
                            ForEach-Object {($_.CounterSamples.CookedValue); If($_.CounterSamples.CookedValue -ge $RequestWarnThreshold){$RequestsWarn = $True};`
                                If ($_.CounterSamples.CookedValue -ge $RequestAlertThreshold){$RequestsAlert = $True}; If ($RequestsAlert = $True){$RequestsWarn = $False}}
            
            $ConAlert = $false
            $ConWarn = $false
            $Connections = Get-Counter -ComputerName $ServerName -Counter "\web service(_total)\current connections" -ErrorAction SilentlyContinue |
                               ForEach-Object {($_.CounterSamples.CookedValue); If($_.CounterSamples.CookedValue -ge $ConnWarnThreshold){$ConWarn = $True};`
                                If ($_.CounterSamples.CookedValue -ge $ConnAlertThreshold){$ConAlert = $True}; If ($ConAlert = $True){$ConWarn = $False}}
                                     

	    }
	
        ## Put the results together
        $Result += New-Object PSObject -Property @{
	        ServerName = $ServerName
		    IPV4Address = $IPv4Address
		    Status = $PingStatus
            Uptime = $Uptime
            CpuUsage = $CpuUsage
            CpuWarn = $CpuWarn
            CpuAlert = $CpuAlert
		    MemUsage = $MemUsage
            MemWarn = $MemWarn
            MemAlert = $MemAlert
            Service = $Service
            ServiceWarn = $ServiceWarn
            ServiceAlert = $ServiceAlert
            Requests = $Requests
            RequestsWarn = $RequestsWarn
            RequestsAlert = $RequestsAlert
            Connections = $Connections
            ConWarn = $ConWarn
            ConAlert = $ConAlert            
	    }

        ## Clear the variables after obtaining and storing the results so offline servers don't have duplicate info.
        #Clear-Variable IPv4Address
        #Clear-Variable Uptime
        #Clear-Variable MemUsage
        #Clear-Variable CpuUsage
    }

$sort = @"
    <script>
        function sortTable(n) {
        var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
        table = document.getElementById("header");
        switching = true;
        // Set the sorting direction to ascending:
        dir = "asc"; 
        /* Make a loop that will continue until no switching has been done: */
        while (switching) {
        // Start by saying: no switching is done:
        switching = false;
        rows = table.getElementsByTagName("TR");
        /* Loop through all table rows (except the first, which contains table headers): */
        for (i = 1; i < (rows.length - 1); i++) {
        // Start by saying there should be no switching:
        shouldSwitch = false;
        /* Get the two elements you want to compare, one from current row and one from the next: */
        x = rows[i].getElementsByTagName("TD")[n];
        y = rows[i + 1].getElementsByTagName("TD")[n];
        /* Check if the two rows should switch place, based on the direction, asc or desc: */
        if (dir == "asc") {
        if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
        // If so, mark as a switch and break the loop:
        shouldSwitch= true;
        break;
        }
        } else if (dir == "desc") {
        if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
        // If so, mark as a switch and break the loop:
        shouldSwitch= true;
        break;
      }
    }
  }
        if (shouldSwitch) {
        /* If a switch has been marked, make the switch and mark that a switch has been done: */
        rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
        switching = true;
        // Each time a switch is done, increase this count by 1:
        switchcount ++; 
        } else {
        /* If no switching has been done AND the direction is "asc", set the direction to "desc" and run the while loop again. */
        if (switchcount == 0 && dir == "asc") {
        dir = "desc";
        switching = true;
      }
    }
  }
}
</script>
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
                </head>
"@

$HTML += @"
                <html><body>
                $sort
                <table border=1 cellpadding=0 cellspacing=0 id=header>
                <thead>
                <tr>
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>Server</font></b></th>
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>IP</font></b></th>
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>Status</font></b></th>
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>DLS Memory Usage</font></b></th>
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>CPU Usage</font></b></th>
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>Memory Usage</font></b></th>
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>Queuing</font></b></th>
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>Total Connections</font></b></th>                
                <th onclick="sortTable(0)"><b><font color=#e6e6e6>Uptime</font></b></th>
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
        ## Symbol &#128543 = Worried Face
        ## Symbol &#128552 = fearful face
        ## Symbol &#128561 = face screaming in fear
        ForEach($Entry in $Result)
        {   
            #Server Name
            If ($Entry.Status -eq $True)
            {
                $HTML += "<td><font color=#00e600>$($Entry.ServerName)</font></td>"
            }

            Else
            {
                $HTML += "<td><blink><font color=#FF4D4D>&#10008 $($Entry.ServerName)</font></blink></td>"
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
                If ($Entry.ServiceWarn -eq $true)
                    {   # Warning! Service Mem Usage is equal to or above warning threshold!
                        $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.Service)</font></td>"
                    }
                Else
                    {   # Everything is OK!
                        $HTML += "<td><font color=#00e600>&#10004 $($Entry.Service)</font></td>"
                    }
                If ($Entry.ServiceAlert -eq $true)
                    {   # Alert!! Service Mem usage is equal to or above the alert threshold!! 
                        $HTML +="<td><font color=#FF4D4D>&#128561 $($Entry.Service)</font></td>"
                    }
                Else
                    {   # Service not running?!
                        $HTML += "<td><font color=#FF4D4D>&#10008 Check DataLayerService</font></td>"
                    }
            }

            #CPU Usage
            If ($Entry.CpuUsage -ne $null)
            {
                If ($Entry.CpuWarn -eq $True)
                    {   # Warning! CPU usage is equal to or above warning threshold!
                        $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.CpuUsage)</font></td>"
                    }
                Else
                    {   # Everything is OK!
                        $HTML += "<td><font color=#00e600>&#10004 $($Entry.CpuUsage)</font></td>"
                    }
             If ($Entry.CpuAlert -eq $true)
                    {   # Alert!! CPU usage is equal to or above the alert threshold!!
                        $HTML += "<td><font color=#FF4D4D>&#128561 $($Entry.CpuUsage)</font></td>"
                    }
            }

            # Memory Usage
            If ($Entry.MemUsage -ne $null)
            {
                If ($Entry.MemWarn -eq $True)
                    {   # Warning! Mem usage is equal to or above the warning threshold!
                        $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.MemUsage)</font></td>"
                    }
                Else
                    {   # Everything is OK!
                        $HTML += "<td><font color=#00e600>&#10004 $($Entry.MemUsage)</font></td>"
                    }
                
                If ($Entry.MemAlert -eq $True)                
                    {   # Alert!! Mem usage is equal to or above the alert threshold!!
                        $HTML += "<td><font color=#FF4D4D>&#128561 $($Entry.MemUsage)</font></td>"
                    }
             }              

            # check ASP.NET Queued Requests
            If ($Entry.Requests -ne $null)
            {
                If ($Entry.RequestsWarn -eq $True)
                    {   # Warning! Web Requests are equal to or above the warning threshold!
                        $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.Requests)</font></td>"
                    }
                Else
                    {   # Everything is OK!
                        $HTML += "<td><font color=#00e600>&#10004 $($Entry.Requests)</font></td>"
                    }
                If ($Entry.RequestsAlert -eq $True)
                    {   # Alert!! Web Requests are equal to or above the warning threshold!!
                        $HTML += "<td><font color=#FF4D4D>&#128561 $($Entry.Requests)</font></td>"
                    }   
                Else
                    {   # No Counter Found!
                        $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                    }
            }
            # check Total Connections
            If ($Entry.Connections -ne $null)
            {
                If ($Entry.ConWarn -eq $True)
                    {   # Warning! Web Connections are equal to or above the warning threshold!
                        $HTML += "<td><font color=#ffff4d>&#9888 $($Entry.Connections)</font></td>"
                    }
                Else
                    {   # Everything is OK!
                        $HTML += "<td><font color=#00e600>&#10004 $($Entry.Connectons)</font></td>"
                    }
                If ($Entry.ConAlert -eq $True)
                    {   # Alert!! Web Connections are equal to or above the warning threshold!!
                        $HTML += "<td><font color=#FF4D4D>&#128561 $($Entry.Connections)</font></td>"   
                    }
                Else
                    {   # No Counter Found!
                        $HTML += "<td><font color=#FF4D4D>&#10008 Counter Does Not Exist</font></td>"
                    }
            }

             #Uptime Status
            If ($Entry.Status -eq $True)
            {
                $HTML += "<td><font color=#00e600>$($Entry.Uptime)</font></td>
                          </tr>"
            }

            Else
            {
                $HTML += "<td><font color=#FF4D4D>&#10008 Offline</font></td>
                          </tr>"
            }

        } }

}
        ## Report the date and time the script ran.
                 
        $HTML += "</table><p><font color=#e6e6e6>Status refreshed on: $(Get-Date -Format G) Script Time " + $TimeToRun.TotalSeconds + " Seconds</font></p></body></html>"
       
        ## Output the HTML file
	    $HTML | Out-File $OutputFile
      

        ## If the refresh time option is configured, wait the specifed number of seconds then loop.
        If ($RefreshTime -ne $null)
        {
            Start-Sleep -Seconds $RefreshTime
        }
    
}

## If the refresh time option is not configured, stop the loop.
Until ($RefreshTime -eq $null)

## End
