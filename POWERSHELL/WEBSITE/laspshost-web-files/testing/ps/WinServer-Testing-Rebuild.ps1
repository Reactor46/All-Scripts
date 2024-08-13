$ServerList = Get-Content "C:\inetpub\wwwroot\testing\ps\Configs\TEST.txt" -ErrorAction SilentlyContinue
$ReportFilePath = "C:\inetpub\wwwroot\testing\reports\SeverHealthSortable.htm"
$Result = @()

Function Get-UpTime
{
    param([string] $LastBootTime)
    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
    "$($Uptime.Days) days $($Uptime.Hours)h $($Uptime.Minutes)m"
}

$TimeToRun=Measure-Command {

ForEach($ComputerName in $ServerList){
    $PingStatus = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet
    
    ## If server responds, get uptime and disk info
        If ($PingStatus)
        {     
        $AVGProc = Get-CimInstance -computername $ComputerName win32_processor -ErrorAction SilentlyContinue | select -ExpandProperty LoadPercentage 
            $log = New-Object psobject -Property @{
            Server = $ComputerName
            AVGProc = ($AVGProc | Measure-Object -Average).Average
            }
        $IPv4Address = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ServerName |
                                Select-Object -Expand IPAddress | Select -First 1 |
                                    Where-Object { ([Net.IPAddress]$_).AddressFamily -eq "InterNetwork" }
                                       
        $OS = Get-CimInstance win32_operatingsystem -computername $ComputerName -ErrorAction SilentlyContinue | Select-Object @{Name = "MemoryUsage"; Expression = {“{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}
        $Service = Get-Process -ComputerName $ComputerName -Name DataLayerService -ErrorAction SilentlyContinue | Select-Object  @{Name = "MemoryUsage"; Expression = {“{0:F0}” -f (($_.WS))}} | Sort -Property WS
        $Requests = Get-Counter -ComputerName $ComputerName -Counter "\asp.net\requests queued" -ErrorAction SilentlyContinue
        $CurrentTotalConnect = Get-Counter -ComputerName $ComputerName -Counter "\web service(_total)\current connections" -ErrorAction SilentlyContinue
 
        $Result += [PSCustomObject] @{ 
        ServerName = "$ComputerName"
        IPV4Address = $IPv4Address
        Status = $PingStatus
        CPULoad = $log.AVGProc
        MemLoad = $OS.MemoryUsage
        ContosoDLS = $Service.MemoryUsage
        Requests = $Requests.CounterSamples.CookedValue
        TotalCon = $CurrentTotalConnect.CounterSamples.CookedValue
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

        ## If there is a result put the HTML file together.
            If ($Result -ne $null){

                Foreach($Entry in $Result) 
            { 
          #convert raw data to percentages
          $CPUAsPercent = "$($Entry.CPULoad)%"
          $MemAsPercent = "$($Entry.MemLoad)%"
          $DLS = "$(($Entry.ContosoDLS)/1MB)"
          $ReqTotal = "$($Entry.Requests)"
          $ConTotal = "$($Entry.TotalCon)"

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
           
        
          if(($Entry.ContosoDLS)/1MB -gt 1800)
          {
              $HTML += "<TD $Style1>$([math]::Round($DLS)) MB</TD>"
          }
          elseif((($Entry.ContosoDLS)/1MB -ge 1000) -and (($Entry.ContosoDLS)/1MB -lt 1799))
          {
              $HTML += "<TD $Style2>$([math]::Round($DLS)) MB</TD>"
          }
          else
          {
              $HTML += "<TD $Style3>$([math]::Round($DLS)) MB</TD>"
          }

          # check CPU load
          if(($Entry.CPULoad) -ge 80) 
          {
              $HTML += "<TD $Style1>$($CPUAsPercent)</TD>"
          } 
          elseif((($Entry.CPULoad) -ge 70) -and (($Entry.CPULoad) -lt 80))
          {
              $HTML += "<TD $Style2>$($CPUAsPercent)</TD>"
          }
          else
          {
              $HTML += "<TD $Style3>$($CPUAsPercent)</TD>" 
          }

          # check RAM load
          if(($Entry.MemLoad) -ge 80)
          {
              $HTML += "<TD $Style1>$($MemAsPercent)</TD>"
          }
          elseif((($Entry.MemLoad) -ge 70) -and (($Entry.MemLoad) -lt 80))
          {
              $HTML += "<TD $Style2>$($MemAsPercent)</TD>"
          }
          else
          {
              $HTML += "<TD $Style3>$($MemAsPercent)</TD>"
          }
          
         # check ASP.NET Queued Requests
          if(($Entry.Request) -ge 251)
          {
              $HTML += "<TD $Style1>$($ReqTotal)</TD>"
          }
          elseif((($Entry.Requests) -ge 100) -and (($Entry.Requests) -lt 250))
          {
              $HTML += "<TD $Style2>$($ReqTotal)</TD>"
          }
          else
          {
              $HTML += "<TD $Style3>$($ReqTotal)</TD>"
          }

          # check Total Connections
          if(($Entry.TotalCon) -ge 500)
          {
              $HTML += "<TD $Style1>$($ConTotal)</TD>"
          }
          elseif((($Entry.Requests) -ge 300) -and (($Entry.TotalCon) -lt 499))
          {
              $HTML += "<TD $Style2>$($ConTotal)</TD>"
          }
          else
          {
              $HTML += "<TD $Style3>$($ConTotal)</TD>"
          }              
         
          $HTML += "</TR>"
    }}

    $HTML += "</Table></BODY></HTML>" + "</table><Div id=""footer"">Script Time " + $TimeToRun.TotalSeconds + " Seconds</div></div></div>"
} 
 
$HTML | Out-File $ReportFilePath