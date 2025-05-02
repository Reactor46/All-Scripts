$ServerList = Get-Content "C:\Scripts\Repository\jbattista\Web\reports\ServerHealth\Configs\Servers.txt" -ErrorAction SilentlyContinue
$ReportFilePath = "C:\Scripts\Repository\jbattista\Web\reports\Server_Health.htm"
$Result = @()

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
	refreshpage(120, "countdown") // refreshpage(duration_in_seconds, id_of_element_to_show_result)
}

</script>

<div>Next <a href="javascript:window.location.reload()">refresh</a> in <b id="countdown"></b> seconds</div>
"@

$TimeToRun=Measure-Command {

ForEach($ComputerName in $ServerList){
$AVGProc = Get-WmiObject -computername $ComputerName win32_processor -ErrorAction SilentlyContinue | select -ExpandProperty LoadPercentage 
    $log = New-Object psobject -Property @{
    Server = $ComputerName
    AVGProc = ($AVGProc | Measure-Object -Average).Average
    }   
$OS = Get-CimInstance win32_operatingsystem -computername $ComputerName -ErrorAction SilentlyContinue | Select-Object @{Name = "MemoryUsage"; Expression = {“{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}
$Service = Get-Process -ComputerName $ComputerName -Name DataLayerService -ErrorAction SilentlyContinue | Select-Object  @{Name = "MemoryUsage"; Expression = {“{0:F0}” -f (($_.WS))}} | Sort -Property WS
$Requests = Get-Counter -ComputerName $ComputerName -Counter "\asp.net\requests queued" -ErrorAction SilentlyContinue
$CurrentTotalConnect = Get-Counter -ComputerName $ComputerName -Counter "\web service(_total)\current connections" -ErrorAction SilentlyContinue
 
$Result += [PSCustomObject] @{ 
        ServerName = "$ComputerName"
        CPULoad = $log.AVGProc
        MemLoad = $OS.MemoryUsage
        ContosoDLS = $Service.MemoryUsage
        Requests = $Requests.CounterSamples.CookedValue
        TotalCon = $CurrentTotalConnect.CounterSamples.CookedValue
  }


    $OutputReport = "<HTML><TITLE>Server Health Report</TITLE>
                     <BODY bgcolor=#1557c1>
                     <font color =#FFFFFF>
                     <H2>Server Monitor Report</H2>
                     $timer
                     <Table border=2 cellpadding=4 cellspacing=3>
                     <TR bgcolor=D1D0CE align=center>
                     <TD><B>Server</B></TD>
                     <TD><B>DLS Memory Usage</B></TD>
                     <TD><B>Sys CPU %</B></TD>
                     <TD><B>Sys Mem %</B></TD>
                     <TD><B>Queuing</B></TD>
                     <TD><B>Total Connections</B></TD>
                     </TR>"

$Style1 = @"
style="background-color:Crimson;color:white;text-align:center"
"@
$Style2 = @"
style="background-color:Yellow;color:black;text-align:center"
"@
$Style3 = @"
style="background-color:Lime;color:black;text-align:center"
"@
                          
    Foreach($Entry in $Result) 
    { 
          #convert raw data to percentages
          $CPUAsPercent = "$($Entry.CPULoad)%"
          $MemAsPercent = "$($Entry.MemLoad)%"
          $DLS = "$(($Entry.ContosoDLS)/1MB)"
          $ReqTotal = "$($Entry.Requests)"
          $ConTotal = "$($Entry.TotalCon)"
         

          $OutputReport += "<TR bgcolor=D1D0CE align=left><TD>$($Entry.Servername)</TD>"
          if(($Entry.ContosoDLS)/1MB -gt 1800)
          {
              $OutputReport += "<TD $Style1>$([math]::Round($DLS)) MB</TD>"
          }
          elseif((($Entry.ContosoDLS)/1MB -ge 1000) -and (($Entry.ContosoDLS)/1MB -lt 1799))
          {
              $OutputReport += "<TD $Style2>$([math]::Round($DLS)) MB</TD>"
          }
          else
          {
              $OutputReport += "<TD $Style3>$([math]::Round($DLS)) MB</TD>"
          }

          # check CPU load
          if(($Entry.CPULoad) -ge 80) 
          {
              $OutputReport += "<TD $Style1>$($CPUAsPercent)</TD>"
          } 
          elseif((($Entry.CPULoad) -ge 70) -and (($Entry.CPULoad) -lt 80))
          {
              $OutputReport += "<TD $Style2>$($CPUAsPercent)</TD>"
          }
          else
          {
              $OutputReport += "<TD $Style3>$($CPUAsPercent)</TD>" 
          }

          # check RAM load
          if(($Entry.MemLoad) -ge 80)
          {
              $OutputReport += "<TD $Style1>$($MemAsPercent)</TD>"
          }
          elseif((($Entry.MemLoad) -ge 70) -and (($Entry.MemLoad) -lt 80))
          {
              $OutputReport += "<TD $Style2>$($MemAsPercent)</TD>"
          }
          else
          {
              $OutputReport += "<TD $Style3>$($MemAsPercent)</TD>"
          }
          
         # check ASP.NET Queued Requests
          if(($Entry.Request) -ge 251)
          {
              $OutputReport += "<TD $Style1>$($ReqTotal)</TD>"
          }
          elseif((($Entry.Requests) -ge 100) -and (($Entry.Requests) -lt 250))
          {
              $OutputReport += "<TD $Style2>$($ReqTotal)</TD>"
          }
          else
          {
              $OutputReport += "<TD $Style3>$($ReqTotal)</TD>"
          }

          # check Total Connections
          if(($Entry.TotalCon) -ge 500)
          {
              $OutputReport += "<TD $Style1>$($ConTotal)</TD>"
          }
          elseif((($Entry.Requests) -ge 300) -and (($Entry.TotalCon) -lt 499))
          {
              $OutputReport += "<TD $Style2>$($ConTotal)</TD>"
          }
          else
          {
              $OutputReport += "<TD $Style3>$($ConTotal)</TD>"
          }              
         
          $OutputReport += "</TR>"
    }}

    $OutputReport += "</Table></BODY></HTML>" + "</table><Div id=""footer"">Script Time " + $TimeToRun.TotalSeconds + " Seconds</div></div></div>"
} 
 
$OutputReport | Sort-Object $DLS | out-file $ReportFilePath