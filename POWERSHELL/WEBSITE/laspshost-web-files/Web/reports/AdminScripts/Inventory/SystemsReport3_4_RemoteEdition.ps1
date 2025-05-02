
$version = "3.4 (Remote Edition)" #<<<<<<CHANGE THIS ANYTIME MODIFIED!
#region Note
<# Note: WinRM must be enabled to run the following portions of this script:
    *NonWinServices
    *Installed Software Report
    *Driver Report
If WinRM cannot be enabled, the above listed portions of the script will NOT work. Comment them out with a # if necessary
Enable WinRM by running ' winrm quickconfig ' (no quotes) at an elevated command prompt
By default PowerShell will use the following ports for communication (They are the same ports as WinRM)
see http://blogs.technet.com/b/christwe/archive/2012/06/20/what-port-does-powershell-remoting-use.aspx for more info.
*TCP/5985 = HTTP
*TCP/5986 = HTTPS
Windows Firewall rules required
-"Windows Management Instrumentation" (DCOM-In) and (WMI-In)
-"Remote Event Log Management" (RPC/EPMAP) 
-"File and Printer Sharing" (SMB-In) - or port 445 in.
*Updates:
  - Tested and verified Active Directory integration. See line 206 comments.
#>
#endregion note
#region Information Gathered
<#
The following information is reported upon:
(In HTML Report)
-System Information 
    -System Uptime
    -System Manufacturer
    -System Model
    -OS
    -CPU Name and Bit-ness
    -CPU Cores (Physical & Logical)
    -CPU Usage %
    -Total RAM
    -Free RAM
    -Percent free RAM
  -Disk Info & Performance
    -Physical Disk
      -Model
      -DeviceID
      -Size (GB)
      -Disk# Name
      -AvgDiskQueueLength
      -% Idle Time
      -% Write Time
      -% Read Time
      -Write IOPS
      -Read IOPS      
  - System Processes and Services
    -System Processes - Top 10 Highest Memory Usage
    -System Services - Stopped
    -Startup Entries
    -Non-Windows Services
    -Non-Standard Service Accounts
    -3rd party Scheduled Tasks
-Networking Information
    -Active NICs
      -Hostname
      -NIC Name
      -Network:
      -Link Speed
      -Up Time
      -MAC Address
      -IP Address
      -DHCP Server
      -Subnet Mask
      -Default Gateway
      -DNS Suffix
      -DNS Servers
    -Route Tables    
-Event Logs
    -Event type of either Warning, Error, or Critical within 24 hrs.
-Software Information
    -Installed Programs
    -Installed Updates
    -Loaded 3rd-Party Windows Drivers

 #>
 #endregion Information Gathered             
 #region Change log
 <#
__________________________________________________________
Change Log
   3.4.0 (Remote Edition)
   -Added .NET version detection

   3.3.0 (Remote Edition)
   -Added Disk Status from Win32_Diskdrive query
   
   3.2.0 (Remote Edition)
   -CSS Change
   -Changed Driver Function to return only 3rd party kernel drivers
   -Changed Ping test to increase speed of ping job that checks server's connectivity.
   -Tested and verified Active Directory integration. See line 206 comments
     
   3.1.0 (Remote Edition)
   -Changed Ping test to increase speed of ping job that checks server's connectivity.
   -Changed System Info to its own function, slimming down repetative WMI queries
   -Added BIOS information if machine is physical, left out if virtual
  
   3.0 (Remote Edition)
   -Ran each server queried as jobs, allowing MULTICORE ability to process data, dramatically improving execution time of script
   -Made major sections of the report collapsible for cleaner look, through CSS
   -Wording changes in report
   -Cleaned up html
   -Fixed Date sorting in windows updates listing (which allowed removing duplicate queries)
   -Removed double Windows Update wmi queries for better performance
   -Implemented "Top Processes" Bar chart function and implementation
   -Added Get-MemoryStats function
   -Added Memory Module information section to report
   -Changed Memory table to show limited data when polling a VM, while pulling all data from a physical machine
   -Changed process bar and top process report to use WMI instead of "Get-Process" function, for better accuracy and less buggy results
   -Admin check
   -Added Get-WindowsDriver Function, added running drivers section in report
   -Error handling!!! for random memory leak fix and cleaner output
   -Removed Add-HTMLAttribute, replaced with foreach -replace
   -Embedded Images
  
   2.9.0 (Remote Edition)
   -Optimized some WMI queries, by using WQL to improve performance of script
   -Fixed incorrect heading on physical disk perf counters. read time and write time were reversed
   -Fixed Auto-not-started services section. Start Mode now shows values
   -Fixed Process section... had to sort by workingset64, NOT WS :)
   -Fixed query for logical disk section. Missing freespace object
   -Changed text in system services area
   -Removed null entries from $softwarereport section
   -Fixed html update section bug - Missing 'ComputerName' parameter in variable
  
   2.8.8 (Remote Edition)
   -Fixed memory value errors in "Top Process" sections. Changed to workingset64.
    Now, WS MB column reflects SQL server memory consumption accurately. It also references the corresponding sources in Windows.
  
   2.8.6 (Remote Edition)
   -Changed Pie Chart Functions, Added color coding and legends, changed to donut style graph
   -Minor Tweak in table hover color
   -Exception handling in process owner feature
  
   2.8.5 (Remote Edition)
   -Incorporated custom Get-NicInfo function, unifying NIC information from 2 seperate WMI calls into one table
   -Change Network Adapter HTML layout to better fit new function results
   -Tweaked Driver information function to work on remote systems
   -Changed To Remote Edition
   -Clean up old commented out code
   -(Cancelled: Still commented in for possible future use) Script makes sub-directories per computer for easier organization of files
   -Made all commands that could be made remotely executed to do so
   -Removed single machine use CLI references - "region WindowsCommandOutput"
   -Added error handling into Pie Chart Functions
  
   2.8
   -CSS design change
   -Layout tweaking of HTML tables
   -Added additional NIC information
   -Clean up HTML bugs and spelling errors
   -Added process owner to processes table
   -Changed Update query function to use gwmi instead of wmic
   -Made Update links CLICKABLE :)
   -Remove Firewall settings due to inaccuracy, and over complicated function used. Investigating other options
  
   2.7
   -HTML design changes
   -Changed Active NIC Table headers to a more english friendly format
     
   2.6
   -Fixed CPU Info bug that showed null on some servers
   -Added .cmd Executable to bypass the powershell execution policy steps on other machines
  
   2.5
   -Documented data gathered in script
   -Added Physical disk Information
   -Removed Drive Type from "Logical disks" table
   -Added Physical Disk Performance WMI data
   -Marked firewall settings in Security Settings category
   -Added list installed software
   -Added more detail in driver report, made function based execute
   -Fixed Date output in Installed Software
   -Added 32 bit Programs to program list
   -Added Non-Windows Services section in HTML report
  
   2.0
   -Added Functions Get-RemoteFirewallStatus, Get-RemoteRouteTable, Get-RemoteScheduledTasks
   -Incorporated all functions into report
   -Massive layout change to html and css of report
   -Cleaned up html
   -Specified WindowsCommandOutputRegion, moved driver query into it *Driver query broke, moved back to top.
   -Moved get-winevent filter variables to var&arg section
   -Pull all warnings/errors from System and Application event logs with wevtutil and saved to file in evtx format
   -Pull RSOP with gpresult
   -Tidy up script region markers
  
PowerShell Systems Report
Example usage: .\SystemsReport.ps1 .\list.txt
Remember that list.txt is the file containing a list of Server names to run this against
Parts of script commented out (#), may be turned back on to add multiple computer ability and email functions.
#>
#endregion Change log
#$users = "youremail@yourcompany.com" # List of users to email your report to (separate by comma)
#$fromemail = "youremail@yourcompany.com"
#$server = "yourmailserver.yourcompany.com" #enter your own SMTP server DNS name / IP address here
#$list = $args[0] #This accepts the argument you add to your scheduled task for the list of servers. i.e. list.txt
#region Admin Check
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] “Administrator”))
{
    Write-Warning “Administrator rights are required to run this script! Please re-run this script as an Administrator!”
    Start-Sleep -Seconds 5
    Break
}
#endregion Admin Check
$computers = get-content list.txt #comment this line out if using AD
<# AD Integration... Uncomment to gather all AD enabled computer accounts and run this report. Replaces list.
Import-Module ActiveDirectory
$computers = Get-ADComputer -filter {enabled -eq $True} -Properties Name|Sort Name |Select -expand Name
#>
cls
""
Write-Host -ForegroundColor Yellow "Checking Computers for connectivity before running script..."
Write-Host -ForegroundColor Yellow "See Errorlog.txt for servers that failed to connect"
Start-Sleep -Seconds 2
""
#region Remote Check
$CompPlaceHolder = @()
foreach ($computer in $computers) {
   
     
try
{
    Write-Host -ForegroundColor Yellow “Verifying Remote RPC/WMI Connectivity on" $computer "...”
    $TestConnect = Test-Connection -ComputerName $computer -Count 1 -WarningAction Stop -EA Stop
    $WMIConnect= Get-WmiObject Win32_ComputerSystem -computername $computer -EA Stop
    Write-Host -ForegroundColor Green "Remote RPC/WMI connection successful!"
    ""
}
       
catch                
      {
     
      Write-Host -ForegroundColor Red 'Remote RPC/WMI on server:' $computer 'is unreachable. See Errorlog.txt'
      $LogMsgRpc = ("Error Connecting to RPC/WMI on " + $computer + " - "+ (Get-Date))
      Write-Output $LogMsgRpc  |Out-File -FilePath .\Errorlog.txt -Append
     
     
          }
                   
try
{
   
    Write-Host ""
    Write-Host -ForegroundColor Yellow “Verifying Remote Event Log Connectivity on" $computer "...”
    $EvtTestconnect = Test-Connection -ComputerName $computer -Count 1 -WarningAction Stop -EA Stop
    $EVTConnect= Get-WinEvent -ComputerName $computer -LogName application -MaxEvents 1 -ErrorAction Stop
    Write-Host -ForegroundColor Green "Remote Event Log connection successful"
    ""
    $Connections = New-Object -Type PSObject -Property @{
            "Name"          = $computer
            }
            $CompPlaceHolder += $Connections
}
catch
{
   
    Write-Host -ForegroundColor Red 'Server' $computer 'is unreachable. Unable to continue. Error connecting to remote event logs.
    Verify Remote Event Log Management is enabled on' $computer '. See Errorlog.txt'
    $LogMsgEvt = ("Error Connecting to Remote Event Logs on " + $computer + " - "+ (Get-Date))
    Write-Output $LogMsgEvt |Out-File -FilePath .\Errorlog.txt -Append
    Write-Output "------------------------------------" |Out-File -FilePath .\Errorlog.txt -Append
    Write-Output "" |Out-File -FilePath .\Errorlog.txt -Append
      
}
}
$computers = $CompPlaceHolder|Select -ExpandProperty Name #Use of $CompPlaceHolder.name doesn't work right in PS 2.0
#endregion Remote Check
#* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
#Scriptblock to run as jobs
""
Write-Host -ForegroundColor Yellow "Connection tests passed. Running jobs in multithreaded mode"
""
$Scriptblock= {
   
    param($computer)
#region Variables and Arguments   
    $Timer=Get-Random -Maximum 300 -Minimum 100
    Start-Sleep -m $timer
[System.Threading.Thread]::CurrentThread.Priority = 'BelowNormal'   
[int]$ProcessNumToFetch = 10
$ListOfAttachments = @()
$Report = @()
$date= Get-Date -format g
$version = "3.1 (Remote Edition)" #<<<<<<CHANGE THIS ANYTIME MODIFIED!
$SysEvtFilter = [xml]@'
<QueryList>
  <Query Id="0" Path="System">
    <Select Path="System">*[System[(Level=1  or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) &lt;= 86400000]]]</Select>
  </Query>
</QueryList>
'@
$AppEvtFilter = [xml]@'
<QueryList>
  <Query Id="0" Path="Application">
    <Select Path="Application">*[System[(Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) &lt;= 86400000]]]</Select>
  </Query>
</QueryList>
'@
$RedBang = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAA2FBMVEX///++MjLnalHTJyDVKy3veGf8p6Tpd2HWYl/UKiPvemncTU/leF3fWTf2kYjZPkDWLjDUXlnwhHXXOjTUSyrYaGf1iH7NS0ncYj/USy/eVjTPJxPMJwDMJwb1hnzYOz7ONAX8o6DVQxnhbG7NTE3QKhfZPkHWVS7SXlfic3XnbFTZU0P+/v7PVle1IiHMR0ieGwD1qqD3xLyzSTLLf3DPNBaQHwL99fTcalDjeHTCb2uOFQ7KJhL21MyGJxC3LhWoHQTvuKnkycKAHgjsrqb01tPmmYLSnpF1j2u3AAAAAXRSTlMAQObYZgAAAKxJREFUeF5lz0WWw0AMRVEXmZkdZmZmbNj/jlpKO6PcyT/1RioJkZyUI5FcB3KUF6LacQXEtvoqxPN1IeYLIXTfg0J6yoBSuh0tKR0oPSIR0yqA8+gCY5kQ0ga6P544KYRuDf38bnC6EMIS2q1mOCGErIzW1xtOBmHYLoKv7xNMewghCTTG2P5wZEwLEjxk6ric88mEu870/9SW0amCjtF6f2bcb4L+GN8f3/8DQgYUGINoSt8AAAAASUVORK5CYII='
$RedBangHTML = "<img src=data:image/png;base64,$($RedBang) alt='Error' style='vertical-align:middle;padding-right:5px;'/>"
$YellowBang = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAACdklEQVR42mVSTU8TURR9JCQsoJKIq0oUookpBkJnphqXtD+AX8BfcMOWD2PbtCaaYCKRIhRFi1q1Frvwk6ihKMFFi9ICg0BRISIFtbQKKe3xvjcd08LizH25953zzr13mNVqZS0tLQcw55TaVLcE1W3GvEtp0+/ZbDb9juCWkCghQOQKTt55XS8ghFxKhS7Co35mxcq6esyhdK8P1wGfmoCPTVgfrkfMLnVT7YBjPVFMbvh8uRHZD8eBGVkgO3lMtDJjlxr2t8qJJarc7s/Ro0C8GblcDvl8HtuTjdjw10K9JGO/C1bc16xDal2+ZiLrJ4A5CblsSohER05hc6wOiasmzDqV1uKZ6Q5EnCebmZdGIK4I7GxNCQf+C7X4GjqNdMhIGzGjRIA+hd6l89/66eXpBiJbgFipQCJA+XcmrPadRJzulmwhelGuWaD+dl/VAlEztdBM0yc3q0H8WHkvBNaCJhKg1ZILPqdpu1Lz/z+IO2TXhpesjzHgLWGcEGbILPViJTKAqSsMSR/lgoTHdO43IuZUXJzLB3eO97U3Wg48Z5rIG44y/H1RjRVfNZY8DNt3KfeQcJ9hz1cuNhInLqMfxPfbe0RTDxGeMk3oGcOftSB2kxEk+hjStyl3j8Cd3GH4df0w6J/xMepdKzzQ7GG0IPSEWljuxZfIoJjB1kCBzIVuEW6WgW+NcfsY0azBX7AZIDyiV4mw0MOg0gxSXv6yTmbIDzExTOEgM2BAZsiAtLcS6aEq7UxIDVYh6anEJiF1w0D3DiHDo4fqhO89dWCB9jPuiQ5ZDXdaFie6LIsiEsY7ZC12KovhLp7T60qhpqiB9rPuf+YeJOq16/K9AAAAAElFTkSuQmCC'
$YellowBangHTML = "<img src=data:image/png;base64,$($YellowBang) alt='Warning' style='vertical-align:middle;padding-right:5px;'/>"
#endregion Variables and Arguments
[System.Threading.Thread]::CurrentThread.Priority = 'BelowNormal'
#region Functions
Function Get-SystemInformation {
    param(
          [string[]]
        $ComputerName=$env:computername
          )
    foreach ($computer in $ComputerName) {
       
        $CompInfo    = Get-WmiObject -Query "Select Manufacturer,Model FROM Win32_ComputerSystem" -computername $computer
        $W32OS       = Get-WmiObject Win32_OperatingSystem -computername $computer
           
          
           $LastBootUpTime = $W32OS.ConvertToDateTime($W32OS.LastBootUpTime)
           $Time = (Get-Date) - $LastBootUpTime
           '{0:00} Days, {1:00} Hours, {2:00} Minutes, {3:00} Seconds' -f $Time.Days, $Time.Hours, $Time.Minutes, $Time.Seconds|Set-Variable Uptime
        $OSPowerInfo = Get-WmiObject -Namespace root\cimv2\power -Class win32_PowerPlan -computername $computer|Where-Object {$_.IsActive -eq "True"}|Select ElementName
        $CpuInfo     = Get-WmiObject -ComputerName $computer -Class win32_processor       
        $CpuName     = $CpuInfo|Select -First 1 -Expand Name
       $CpuAddr     = $CpuInfo|Select -First 1 -Expand DataWidth
        $CpuCores    = $CpuInfo|Select -expand NumberOfCores
        $CpuLogProc  = $CpuInfo|select -Expand NumberOfLogicalProcessors
        $CpuUsage    = $CpuInfo|Select Loadpercentage | Measure-Object LoadPercentage -Average | Select-Object -expand Average
           $BIOS        = Get-WmiObject -computername $computer -Query "SELECT Manufacturer,SMBIOSBIOSVersion,ReleaseDate FROM Win32_Bios"
       
        $SystemInformation = New-Object -Type PSObject -Property @{
            "System Manufacturer" = $CompInfo.Manufacturer
            "System Model"        = $CompInfo.Model
            "BIOS Vendor"         = $BIOS.Manufacturer
            "BIOS Version"        = $BIOS.SMBIOSBIOSVersion
            "BIOS Date"           = ($BIOS.ConvertToDateTime($BIOS.ReleaseDate)).ToShortDateString()
            "CPU Name"            = ($CpuName) + ' x' + ($CpuAddr)
            "Physical CPU Cores"  = $CpuCores|Measure -Sum|Select -ExpandProperty Sum
            "Logical CPU Cores"   = $CpuLogProc|Measure -Sum|select -ExpandProperty Sum
            "System Uptime"       = $Uptime
            "Operating System"    = $W32OS.caption + ( $W32OS.CSDVersion)
            "Power Plan"          = if ($OSPowerInfo.ElementName -notlike "*Performance") {"$($OSPowerInfo.ElementName) (Recommended Plan: High Performance)"}
                                        else {$OSPowerInfo.ElementName}
            "CPU % Usage"         = [string](-join $CpuUsage,"%" ) -replace " ", ""
            "Total RAM (GB)"      = [Math]::Round(($W32OS.TotalVisibleMemorySize/1MB),2)
            "Free RAM (GB)"       = [Math]::Round(($W32OS.FreePhysicalMemory/1MB),2)
            "Free RAM %"          = [Math]::Round((($W32OS.FreePhysicalMemory/$W32OS.TotalVisibleMemorySize) * 100),2)
                }
          
        $SystemInformation
        }
   
}   
Function Create-ProcBarChart {
    param($ComputerName = $env:computername,
          [int32]$ProcessNumber
          )
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
 
# chart object
   try {
   $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
   $chart1.Width = 400
   $chart1.Height = 200
   $chart1.BackColor = [System.Drawing.Color]::White
 
# title
   #[void]$chart1.Titles.Add("Top $ProcessNumber - Memory Usage")
   #$chart1.Titles[0].Font = "Arial,11pt"
   #$chart1.Titles[0].Alignment = "topLeft"
 
# chart area
   $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
   $chartarea.Name = "ChartArea1"
   $Chartarea.AxisX.MajorGrid.LineWidth = 0
   $Chartarea.AxisY.MajorGrid.LineWidth = 0
   $chartarea.AxisY.Title = "Memory (MB)"
   $chartarea.AxisX.Title = "Process Name"
   $chartarea.AxisY.IsLogarithmic = $true
   #$chartarea.AxisY.Interval = 100
   $chartarea.AxisX.Interval = 1
   $chart1.ChartAreas.Add($chartarea)
 
# legend
   $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
   $legend.name = "Legend1"
   $legend.font = "Arial"
   $Legend.docking = "Top"
   $Legend.title = "Processes"
   $Legend.TitleFont ="Arial"
   $legend.alignment = "center"
   $Legend.Istextautofit = $true
   $chart1.Legends.Add($legend)
  
 
# data source
   $datasource = Get-WmiObject -ComputerName $computer -Query "SELECT Description,ProcessID,WorkingSetSize FROM Win32_Process" |Select Description,ProcessID,WorkingSetSize | sort WorkingSetSize -Descending  | Select-Object -First $ProcessNumber
 
# data series
   [void]$chart1.Series.Add("WSMem")
   $chart1.Series["WSMem"].ChartType = "Column"
   $chart1.Series["WSMem"].IsVisibleInLegend = $false
   #$chart1.Series["WSMem"].IsVisibleInLegend = $true
   $chart1.Series["WSMem"].BorderWidth  = 3
   $chart1.Series["WSMem"].chartarea = "ChartArea1"
   $chart1.Series["WSMem"].Legend = "Legend1"
   $chart1.Series["WSMem"].Palette = "SemiTransparent"
   $Chart1.Series["WSMem"]["DrawingStyle"] = "Cylinder"
   $datasource | ForEach-Object {$chart1.Series["WSMem"].Points.addxy( ("PID "+$_.ProcessID +" "+([Math]::Round($_.WorkingSetSize / 1mb))+" MB") , ([Math]::Round($_.WorkingSetSize / 1mb))) }
  
 
# save chart
   $chart1.SaveImage($env:tmp + "\ProcBar-"+$computername +".png","png")
    } 
   catch {
             "Error creating chart. Verify Microsoft Chart Controls for Microsoft .NET Framework 3.5 is installed"
             }
}
Function Create-RAMPieChart {
       param($ComputerName = $env:computername
          )
              
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
       
    #Gather RAM Data
    $SystemInfo = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName  | Select-Object Name, TotalVisibleMemorySize, FreePhysicalMemory
       $TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB
          $FreeRAM = $SystemInfo.FreePhysicalMemory/1MB
          $UsedRAM = $TotalRAM - $FreeRAM
          $RAMPercentFree = ($FreeRAM / $TotalRAM) * 100
          $Free = [Math]::Round($FreeRAM, 2)
          $Used = [Math]::Round($UsedRAM, 2)
         
   
       #Create our chart object
       try {
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart      
       $Chart.Width = 200
       $Chart.Height = 200
       $Chart.Left = 10
       $Chart.Top = 10
   
       #Create a chartarea to draw on and add this to the chart
       $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
       $Chart.ChartAreas.Add($ChartArea)
       [void]$Chart.Series.Add("Data")
       #Add a datapoint for each value specified in the arguments (args)
    Write-Host "Now processing chart value: " + $used
              $datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $Used)
           $datapoint.AxisLabel = "$used GB Used"
        $datapoint.Color ="FireBrick"
           $Chart.Series["Data"].Points.Add($datapoint)
   
    Write-Host "Now processing chart value: " + $Free
              $datapoint1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $Free)
           $datapoint1.AxisLabel = "$Free GB Free"
           $datapoint1.Color ="DodgerBlue"
        $Chart.Series["Data"].Points.Add($datapoint1)
       
   
    $Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut
       $Chart.Series["Data"]["PieLabelStyle"] = "Outside"
       $Chart.Series["Data"]["PieLineColor"] = "Black"
       $Chart.Series["Data"]["PieDrawingStyle"] = "Concave"
       ($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $true
    $Chart.Series["Data"].Font = "Arial"
   
    # Create chart legend
        $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
        $legend.name = "Legend1"
        $legend.font = "Arial"
        $Legend.docking = "Top"
        $Legend.title = "RAM Usage"
        $Legend.TitleFont ="Arial"
        $legend.alignment = "center"
        $Legend.Istextautofit = $true
       
 
        # Add chart legend to chart object
        $chart.legends.add($legend)
       
       #Set the title of the Chart to the current date and time
       #$Title = new-object System.Windows.Forms.DataVisualization.Charting.Title
       #$Chart.Titles.Add($Title)
       #$Chart.Titles[0].Text = "RAM Usage"
    #$Chart.Titles[0].Font = "Arial"
       #Save the chart to a file
       $Chart.SaveImage($env:temp + "\RAM-"+$computername + ".png","png")
   
    }
    catch {
             "Error creating chart. Verify Microsoft Chart Controls for Microsoft .NET Framework 3.5 is installed"
             } 
}
Function Create-CPUPieChart {
       param($ComputerName = $env:computername
          )
              
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
       
    #Get CPU Data
    $CpuInfo = Get-WmiObject -ComputerName $computername -Class win32_processor
    $CpuName = $CpuInfo|Select -First 1 -Expand Name
    $CpuAddr = $CpuInfo|Select -First 1 -Expand DataWidth
    $CpuCores = $CpuInfo|Select -expand NumberOfCores
    $CpuUsage = $CpuInfo|Select Loadpercentage | Measure-Object LoadPercentage -Average | Select-Object -expand Average
    $CpuFree = 100 - [int]$CpuUsage
   
       
    #Create our chart object
       try {
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart      
       $Chart.Width = 200
       $Chart.Height = 200
       $Chart.Left = 10
       $Chart.Top = 10
   
       #Create a chartarea to draw on and add this to the chart
       $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
       $Chart.ChartAreas.Add($ChartArea)
       [void]$Chart.Series.Add("Data")
       #Add a datapoint for each value specified in the arguments (args)
    Write-Host "Now processing chart value: " + $CpuUsage
              $datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $CpuUsage)
           $datapoint.AxisLabel = "$CpuUsage% Used"
        $datapoint.Color ="FireBrick"
           $Chart.Series["Data"].Points.Add($datapoint)
   
    Write-Host "Now processing chart value: " + $CpuFree
              $datapoint1 = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $CpuFree)
           $datapoint1.AxisLabel = "$CPUFree% Free"
           $datapoint1.Color ="DodgerBlue"
        $Chart.Series["Data"].Points.Add($datapoint1)
       
   
    $Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut
       $Chart.Series["Data"]["PieLabelStyle"] = "Outside"
       $Chart.Series["Data"]["PieLineColor"] = "Black"
       $Chart.Series["Data"]["PieDrawingStyle"] = "Concave"
       ($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $true
    $Chart.Series["Data"].Font = "Arial"
   
    # Create chart legend
        $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
        $legend.name = "Legend1"
        $legend.font = "Arial"
        $Legend.docking = "Top"
        $Legend.title = "CPU Usage"
        $Legend.TitleFont ="Arial"
        $legend.alignment = "center"
        $Legend.Istextautofit = $true
       
 
        # Add chart legend to chart object
        $chart.legends.add($legend)
       
       #Save the chart to a file
       $Chart.SaveImage($env:tmp + "\CPU-"+$computername + ".png","png")
    }
    catch {
             "Error creating chart. Verify Microsoft Chart Controls for Microsoft .NET Framework 3.5 is installed"
             }
}
Function Get-RemoteRouteTable
{
    <#
    .SYNOPSIS
       Gathers remote system route entries.
    .DESCRIPTION
       Gathers remote system route entries, including persistent routes. Utilizes multiple runspaces and
       alternate credentials if desired.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information
    .EXAMPLE
       PS > Get-RemoteRouteTable
       <output>
      
       Description
       -----------
       <Placeholder>
    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 2.0
       Version History
       1.0.0 - 08/31/2013
        - Initial release
    #>
    [CmdletBinding()]
    PARAM
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
      
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
       
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
       
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )
    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Route Table: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Route Table: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
       
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
       
        Write-Verbose -Message 'Remote Route Table: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Route Table: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
       
        Write-Verbose -Message 'Remote Route Table: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Route Table: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
           
            try
            {
                Write-Verbose -Message ('Remote Route Table: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }
                # General variables
                $ResultSet = @()
                $PSDateTime = Get-Date
                $RouteType = @('Unknown','Other','Invalid','Direct','Indirect')
                $Routes = @()
               
                #region Routes
                Write-Verbose -Message ('Remote Route Table: Runspace {0}: Route table information' -f $ComputerName)
                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','Routes')
                                        
                # WMI data
                $wmi_routes = Get-WmiObject @WMIHast -Class win32_ip4RouteTable
                $wmi_persistedroutes = Get-WmiObject @WMIHast -Class win32_IP4PersistedRouteTable
                foreach ($iproute in $wmi_routes)
                {
                    $Persistant = $false
                    foreach ($piproute in $wmi_persistedroutes)
                    {
                        if (($iproute.Destination -eq $piproute.Destination) -and
                            ($iproute.Mask -eq $piproute.Mask) -and
                            ($iproute.NextHop -eq $piproute.NextHop))
                        {
                            $Persistant = $true
                        }
                    }
                    $RouteProperty = @{
                        'InterfaceIndex' = $iproute.InterfaceIndex
                        'Destination' = $iproute.Destination
                        'Mask' = $iproute.Mask
                        'NextHop' = $iproute.NextHop
                        'Metric' = $iproute.Metric1
                        'Persistent' = $Persistant
                        'Type' = $RouteType[[int]$iproute.Type]
                    }
                    $Routes += New-Object -TypeName PSObject -Property $RouteProperty
                }
                # Setup the default properties for output
                $ResultObject = New-Object PSObject -Property @{
                                                                'PSComputerName' = $ComputerName
                                                                'ComputerName' = $ComputerName
                                                                'PSDateTime' = $PSDateTime
                                                                'Routes' = $Routes
                                                               }
                $ResultObject.PSObject.TypeNames.Insert(0,'My.RouteTable.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                #endregion Routes
                Write-Output -InputObject $ResultObject
            }
            catch
            {
                Write-Warning -Message ('Remote Route Table: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Route Table: Runspace {0}: End' -f $ComputerName)
        }
 
        Function Get-Result
        {
            [CmdletBinding()]
            Param
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Route Table: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Route Table: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Route Table: Getting info'
                        Status = 'Remote Route Table: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Route Table: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
     END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Route Table: Getting route table information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Route Table: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}
Function Get-RemoteScheduledTasks
{
    <#
    .SYNOPSIS
        Gather scheduled task information from a remote system or systems.
    .DESCRIPTION
        Gather scheduled task information from a remote system or systems. If remote credentials
        are provided PSremoting will be utilized.
    .PARAMETER ComputerName
        Specifies the target computer or computers for data query.
    .PARAMETER UseRemoting
        Override defaults and use PSRemoting. If an alternate credential is specified PSRemoting is assumed.
    .PARAMETER ThrottleLimit
        Specifies the maximum number of systems to inventory simultaneously
    .PARAMETER Timeout
        Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
        Show progress bar information
    .EXAMPLE
        PS > (Get-RemoteScheduledTasks).Tasks |
             Where {(!$_.Hidden) -and ($_.Enabled) -and ($_.NextRunTime -ne 'None')} |
             Select Name,Enabled,NextRunTime,Author
        Name                     Enabled   NextRunTime               Author                      
        ----                     -------   -----------               ------                  
        Adobe Flash Player Upd...True      10/4/2013 10:24:00 PM     Adobe
      
        Description
        -----------
        Gathers all scheduled tasks then filters out all which are enabled, has a next
        run time, and is not hidden and displays the result in a table.
    .EXAMPLE
        PS > $cred = Get-Credential
        PS > $Servers = @('SERVER1','SERVER2')
        PS > $a = Get-RemoteScheduledTasks -Credential $cred -ComputerName $Servers
        Description
        -----------
        Using an alternate credential (and thus PSremoting), $a gets assigned all of the
        scheduled tasks from SERVER1 and SERVER2.
       
    .NOTES
        Author: Zachary Loeber
        Site: http://www.the-little-things.net/
        Requires: Powershell 2.0
        Version History
        1.0.0 - 10/04/2013
        - Initial release
       
        Note:
       
        I used code from a few sources to create this script;
            - http://p0w3rsh3ll.wordpress.com/2012/10/22/working-with-scheduled-tasks/
            - http://gallery.technet.microsoft.com/Get-Scheduled-tasks-from-3a377294
       
        I was unable to find several of the Task last exit codes. A good number of them
        from the following source have been included tough;
            - http://msdn.microsoft.com/en-us/library/windows/desktop/aa383604(v=vs.85).aspx
    #>
    [cmdletbinding()]
    PARAM
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        [Parameter(HelpMessage="Override defaults and use PSRemoting. If an alternate credential is specified PSRemoting is assumed.")]
        [switch]
        $UseRemoting,
       
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
        [Parameter(HelpMessage="Display progress of function")]
        [switch]
        $ShowProgress,
       
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
       
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )
    BEGIN
    {
        $ProcessWithPSRemoting = $UseRemoting
        $ComputerNames = @()
       
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Scheduled Tasks: Creating local hostname list'
       
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Scheduled Tasks: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
       
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
            $ProcessWithPSRemoting = $true
        }
       
        Write-Verbose -Message 'Scheduled Tasks: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Scheduled Tasks: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
       
        Write-Verbose -Message 'Scheduled Tasks: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Scheduled Tasks: Defining background runspaces scriptblock'
        $ScriptBlock =
        {
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID,
               
                [Parameter()]
                [switch]
                $UseRemoting
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            $GetScheduledTask = {
                param(
                     $computername = "localhost"
                )
                Function Get-TaskSubFolders
                {
                    param(                      
                        [string]$folder = '\',
                        [switch]$recurse
                    )
                    $folder
                    if ($recurse)
                    {
                        $TaskService.GetFolder($folder).GetFolders(0) |
                        ForEach-Object {
                            Get-TaskSubFolders $_.Path -Recurse
                        }
                    }
                    else
                    {
                        $TaskService.GetFolder($folder).GetFolders(0)
                    }
                }
                try
                {
                     $TaskService = new-object -com("Schedule.Service")
                    $TaskService.connect($ComputerName)
                    $AllFolders = Get-TaskSubFolders -Recurse
                    $TaskResults = @()
                    foreach ($Folder in $AllFolders)
                    {
                        $TaskService.GetFolder($Folder).GetTasks(1) |
                        Foreach-Object {
                            switch ([int]$_.State)
                            {
                                0 { $State = 'Unknown'}
                                1 { $State = 'Disabled'}
                                2 { $State = 'Queued'}
                                3 { $State = 'Ready'}
                                4 { $State = 'Running'}
                                default {$State = $_ }
                            }
                           
                            switch ($_.NextRunTime)
                            {
                                (Get-Date -Year 1899 -Month 12 -Day 30 -Minute 00 -Hour 00 -Second 00) {$NextRunTime = "None"}
                                default {$NextRunTime = $_}
                            }
                            
                            switch ($_.LastRunTime)
                            {
                                (Get-Date -Year 1899 -Month 12 -Day 30 -Minute 00 -Hour 00 -Second 00) {$LastRunTime = "Never"}
                                default {$LastRunTime = $_}
                            }
                            switch (([xml]$_.XML).Task.RegistrationInfo.Author)
                            {
                                '$(@%ProgramFiles%\Windows Media Player\wmpnscfg.exe,-1001)'   { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\acproxy.dll,-101)'                   { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\aepdu.dll,-701)'                     { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\aitagent.exe,-701)'                  { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\appidsvc.dll,-201)'                  { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\appidsvc.dll,-301)'                  { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\System32\AuxiliaryDisplayServices.dll,-1001)' { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\bfe.dll,-2001)'                      { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\BthUdTask.exe,-1002)'                { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\cscui.dll,-5001)'                    { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\System32\DFDTS.dll,-101)'                     { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\dimsjob.dll,-101)'                   { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\dps.dll,-600)'                       { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\drivers\tcpip.sys,-10000)'           { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\defragsvc.dll,-801)'                 { $Author = 'Microsoft Corporation'}
                                '$(@%systemRoot%\system32\energy.dll,-103)'                    { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\HotStartUserAgent.dll,-502)'         { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\kernelceip.dll,-600)'                { $Author = 'Microsoft Corporation'}
                                '$(@%systemRoot%\System32\lpremove.exe,-100)'                  { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\memdiag.dll,-230)'                   { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\mscms.dll,-201)'                     { $Author = 'Microsoft Corporation'}
                                '$(@%systemRoot%\System32\msdrm.dll,-6001)'                    { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\msra.exe,-686)'                      { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\nettrace.dll,-6911)'                 { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\osppc.dll,-200)'                     { $Author = 'Microsoft Corporation'}
                                '$(@%systemRoot%\System32\perftrack.dll,-2003)'                { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\PortableDeviceApi.dll,-102)'         { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\profsvc,-500)'                       { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\RacEngn.dll,-501)'                   { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\rasmbmgr.dll,-201)'                  { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\regidle.dll,-600)'                   { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\sdclt.exe,-2193)'                    { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\sdiagschd.dll,-101)'                 { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\sppc.dll,-200)'                      { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\srrstr.dll,-321)'                    { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\upnphost.dll,-215)'                  { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\usbceip.dll,-600)'                   { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\w32time.dll,-202)'                   { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\wdc.dll,-10041)'                     { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\wer.dll,-293)'                       { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\System32\wpcmig.dll,-301)'                    { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\System32\wpcumi.dll,-301)'                    { $Author = 'Microsoft Corporation'}
                                '$(@%systemroot%\system32\winsatapi.dll,-112)'                 { $Author = 'Microsoft Corporation'}
                                '$(@%SystemRoot%\system32\wat\WatUX.exe,-702)'                 { $Author = 'Microsoft Corporation'}
                                default {$Author = $_ }                                  
                            }
                            switch (([xml]$_.XML).Task.RegistrationInfo.Date)
                            {
                                ''      {$Created = 'Unknown'}
                                default {$Created =  Get-Date -Date $_ }                               
                            }
                            Switch (([xml]$_.XML).Task.Settings.Hidden)
                            {
                                false { $Hidden = $false}
                                true  { $Hidden = $true }
                                default { $Hidden = $false}
                            }
                            Switch (([xml]$_.xml).Task.Principals.Principal.UserID)
                            {
                                'S-1-5-18' {$userid = 'Local System'}
                                'S-1-5-19' {$userid = 'Local Service'}
                                'S-1-5-20' {$userid = 'Network Service'}
                                default    {$userid = $_ }
                            }
                            Switch ($_.lasttaskresult)
                            {
                                '0' { $LastTaskDetails = 'The operation completed successfully.' }
                                '1' { $LastTaskDetails = 'Incorrect function called or unknown function called.' }
                                '2' { $LastTaskDetails = 'File not found.' }
                                '10' { $LastTaskDetails = 'The environment is incorrect.' }
                                '267008' { $LastTaskDetails = 'Task is ready to run at its next scheduled time.' }
                                '267009' { $LastTaskDetails = 'Task is currently running.' }
                                '267010' { $LastTaskDetails = 'The task will not run at the scheduled times because it has been disabled.' }
                                '267011' { $LastTaskDetails = 'Task has not yet run.' }
                                '267012' { $LastTaskDetails = 'There are no more runs scheduled for this task.' }
                                '267013' { $LastTaskDetails = 'One or more of the properties that are needed to run this task on a schedule have not been set.' }
                                '267014' { $LastTaskDetails = 'The last run of the task was terminated by the user.' }
                                '267015' { $LastTaskDetails = 'Either the task has no triggers or the existing triggers are disabled or not set.' }
                                '2147750671' { $LastTaskDetails = 'Credentials became corrupted.' }
                                '2147750687' { $LastTaskDetails = 'An instance of this task is already running.' }
                                '2147943645' { $LastTaskDetails = 'The service is not available (is "Run only when an user is logged on" checked?).' }
                                '3221225786' { $LastTaskDetails = 'The application terminated as a result of a CTRL+C.' }
                                '3228369022' { $LastTaskDetails = 'Unknown software exception.' }
                                default    {$LastTaskDetails = $_ }
                            }
                            $TaskProps = @{
                                'Name' = $_.name
                                'Path' = $_.path
                                'State' = $State
                                'Created' = $Created
                                'Enabled' = $_.enabled
                                'Hidden' = $Hidden
                                'LastRunTime' = $LastRunTime
                                'LastTaskResult' = $_.lasttaskresult
                                'LastTaskDetails' = $LastTaskDetails
                                'NumberOfMissedRuns' = $_.numberofmissedruns
                                'NextRunTime' = $NextRunTime
                                'Author' =  $Author
                                'UserId' = $UserID
                                'Description' = ([xml]$_.xml).Task.RegistrationInfo.Description
                            }
                             $TaskResults += New-Object PSCustomObject -Property $TaskProps
                        }
                    }
                    Write-Output -InputObject $TaskResults
                }
                catch
                {
                    Write-Warning -Message ('Scheduled Tasks: {0}: {1}' -f $ComputerName, $_.Exception.Message)
                }
            }
            try
            {
                Write-Verbose -Message ('Scheduled Tasks: Runspace {0}: Start' -f $ComputerName)
                $RemoteSplat = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                $ProcessWithPSRemoting = $UseRemoting
                if (($LocalHost -notcontains $ComputerName) -and
                    ($Credential -ne [System.Management.Automation.PSCredential]::Empty))
                {
                    $RemoteSplat.Credential = $Credential
                    $ProcessWithPSRemoting = $true
                }
                Write-Verbose -Message ('Scheduled Tasks: Runspace {0}: information' -f $ComputerName)
                $PSDateTime = Get-Date
                $defaultProperties    = @('ComputerName','Tasks')
              if ($ProcessWithPSRemoting)
                {
                    Write-Verbose -Message ('Scheduled Tasks: Using PSremoting on {0}' -f $ComputerName)
                    $Results = @(Invoke-Command  @RemoteSplat `
                                                 -ScriptBlock  $GetScheduledTask `
                                                 -ArgumentList 'localhost')
                    $PSConnection = 'PSRemoting'
                }
                else
                {
                    Write-Verbose -Message ('Scheduled Tasks: Directly connecting to {0}' -f $ComputerName)
                    $Results = @(&$GetScheduledTask -ComputerName $ComputerName)
                    $PSConnection = 'Direct'
                }
               
                $ResultProperty = @{
                    'PSComputerName'= $ComputerName
                    'PSDateTime'    = $PSDateTime
                    'PSConnection'  = $PSConnection
                    'ComputerName'  = $ComputerName
                    'Tasks'         = $Results                   
                }
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
               
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.ScheduledTask.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                Write-Output -InputObject $ResultObject
            }           
            catch
            {
                Write-Warning -Message ('Scheduled Tasks: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Scheduled Tasks: Runspace {0}: End' -f $ComputerName)
        }
       
        Function Get-Result
        {
            [CmdletBinding()]
            Param
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers.($runspace.ID)
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Scheduled Tasks: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Scheduled Tasks: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Scheduled Tasks: Getting info'
                        Status = 'Scheduled Tasks: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        $ComputerNames += $ComputerName
    }
    END
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('UseRemoting',$UseRemoting)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Scheduled Tasks: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
            })
           Get-Result
        }
       
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Scheduled Tasks: Getting share session information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Scheduled Tasks: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}
Function Get-Startup {
param(
    [string]$ComputerName=$env:COMPUTERNAME
    )
$GetStartup = gwmi -ComputerName $ComputerName Win32_StartupCommand |
Select Name, Command, Location, User
Return $GetStartup
}
Function Get-MSHotfix
    {
    param(
    [string]$ComputerName=$env:COMPUTERNAME
    )
$GetMSHotFix = gwmi -ComputerName $ComputerName -Query "SELECT HotFixID,Caption,Description,InstalledOn,InstalledBy FROM win32_quickfixengineering" |
Sort InstalledOn -Descending|
Select HotFixID, Caption, Description, InstalledOn, InstalledBy
Return $GetMSHotFix
}
Function Get-NicInfo{
param([string]$ComputerName=$env:COMPUTERNAME)
foreach ($Adapter in (gwmi Win32_NetworkAdapter -ComputerName $ComputerName -Filter "NetEnabled='True'")) {
$Config = gwmi Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter "Index = '$($Adapter.Index)'"
$Obj= New-Object -Type PSObject -Property @{
        Hostname              = $Adapter.SystemName
        Name                  = $Adapter.name
        "Network"             = $Adapter.NetConnectionID
        "MAC Address"         = $Config.MACAddress
        "IP Address"          = $Config.IPAddress -join "; "
        "DHCP Server"         = if ($Config.DHCPServer -eq $null){"DHCP Disabled"}
                                        else {$Config.DHCPServer}
        "DHCP Enabled"        = $Config.DHCPEnabled
        "DHCP Lease Obtained" = if ($Config.DHCPLeaseObtained -eq $null){"DHCP Disabled"}
                                        else {[management.managementDateTimeConverter]::ToDateTime($Config.DHCPLeaseObtained)}
        "DHCP Lease Expires"  = if ($Config.DHCPLeaseExpires -eq $null){"DHCP Disabled"}
                                        else {[management.managementDateTimeConverter]::ToDateTime($Config.DHCPLeaseExpires)}
        "Subnet Mask"         = $Config.IPSubnet -join "; "
        "Default Gateway"     = $Config.DefaultIPGateway -join "; "
        "DNS Suffix"          = $Config.DNSDomain
        "DNS Servers"         = $Config.DNSServerSearchOrder -join "; "
        "Up Time"             = [management.managementDateTimeConverter]::ToDateTime($Adapter.TimeOfLastReset)
        "Link Speed"          = if ($Adapter.Speed -gt 999999999) {
                                    -join (($Adapter.Speed/1000000000)," ","Gb/s")
                                }
               
                                else {         
                                    -join (($Adapter.Speed/1000000)," ","Mb/s")
                                }
         
    }
$Obj|Select "Hostname",
"Name",
"Network",
"Link Speed",
"Up Time",
"MAC Address",
"IP Address",
"DHCP Server",
"DHCP Enabled",
"DHCP Lease Obtained",
"DHCP Lease Expires",
"Subnet Mask",
"Default Gateway",
"DNS Suffix",
"DNS Servers"
}
}
Function Get-MemoryStats {
param(
    [string]$ComputerName=$env:COMPUTERNAME
    )
        function get-WmiMemoryFormFactor {
        param ([uint16] $char)
        If ($char -ge 0 -and  $char  -le 22) {
        switch ($char) {
        0     {"Unknown"}
        1     {"Other"}
        2     {"SiP"}
        3     {"DIP"}
        4     {"ZIP"}
        5     {"SOJ"}
        6     {"Proprietary"}
        7     {"SIMM"}
        8     {"DIMM"}
        9     {"TSOPO"}
        10     {"PGA"}
        11     {"RIM"}
        12     {"SODIMM"}
        13     {"SRIMM"}
        14     {"SMD"}
        15     {"SSMP"}
        16     {"QFP"}
        17     {"TQFP"}
        18     {"SOIC"}
        19     {"LCC"}
        20     {"PLCC"}
        21     {"FPGA"}
        22     {"LGA"}
        }
        }
       
        else {"{0} - undefined value" -f $char
        }
       
        Return
        }
       
        # Helper function to return memory Interleave  Position
       
        function get-WmiInterleavePosition {
        param ([uint32] $char)
       
        If ($char -ge 0 -and  $char -le 2) {
       
        switch ($char) {
        0     {"Non-Interleaved"}
        1     {"First Position"}
        2     {"Second Position"}
        }
        }
       
        else {"{0} - undefined value" -f $char
        }
       
        Return
        }
       
       
        # Helper function to return Memory Type
        function get-WmiMemoryType {
        param ([uint16] $char)
       
        If ($char -ge 0 -and  $char  -le 24) {
       
        switch ($char) {
        0     {"Unknown"}
        1     {"Other"}
        2     {"DRAM"}
        3     {"Synchronous DRAM"}
        4     {"Cache DRAM"}
        5     {"EDO"}
        6     {"EDRAM"}
        7     {"VRAM"}
        8     {"SRAM"}
        9     {"ROM"}
        10     {"ROM"}
        11     {"FLASH"}
        12     {"EEPROM"}
        13     {"FEPROM"}
        14     {"EPROM"}
        15     {"CDRAM"}
        16     {"3DRAM"}
        17     {"SDRAM"}
        18     {"SGRAM"}
        19     {"RDRAM"}
        20     {"DDR"}
        21     {"DDR2"}
        22     {"DDR2 FB-DIMM"}
        24     {"DDR3"}
        }
       
        }
       
        else {"{0} - undefined value" -f $char
        }
       
        Return
        }
       
       
        # Get the object
        $memory = Get-WMIObject -ComputerName $ComputerName Win32_PhysicalMemory
       
       
       
        Foreach ($stick in $memory) {
       
        # Do some conversions
        $cap=$stick.capacity/1gb
        $ff=get-WmiMemoryFormFactor($stick.FormFactor)
        $ilp=get-WmiInterleavePosition($stick.InterleavePosition)
        $mt=get-WMIMemoryType($stick.MemoryType)
       
        # print details of each stick
        $object = New-Object -type psobject -Property @{
       
        "Bank"                = $stick.DeviceLocator              
        "Capacity (GB)"       = $cap
        "Data Width"          = $stick.DataWidth
        "Description"         = $stick.Description
        "Form Factor"         = $ff
        "InterleaveDataDepth" = $stick.InterleaveDataDepth
        "InterleavePosition"  = $ilp
        "Manufacturer"        = if ($stick.Manufacturer -eq $null) {
                                        "Null - Possibly VM"}
                                        else {$stick.Manufacturer}
        "Memory Type"         = $mt
        "Part Number"         = $stick.PartNumber
        "Serial Number"       = $stick.SerialNumber
        "Speed"               = $stick.Speed
       
        }
        $object #|select Bank,Manufacturer,"Capacity (GB)","Speed","Part Number","Serial Number","Data Width","Form Factor","Memory Type"
        }
       
    }
Function Get-WindowsDriver{
<#
    .SYNOPSIS
        Gather running driver information from a remote or local system.
       
    .PARAMETER ComputerName
        Specifies the target computer for data query.
   
    .EXAMPLE
        PS > Get-WindowsDriver -ComputerName localhost |select Name,Version,Date,Path,Type,State -First 1|ft
            
        Name                          Version                              Date                                 Path                                 Type                                 State                              
        ----                          -------                              ----                                 ----                                 ----                                 -----                              
        Microsoft ACPI Driver         6.1.7600.16385 (win7_rtm.090713-1... 11/20/2010 10:23:47 PM               C:\Windows\system32\drivers\ACPI.sys Kernel Driver                        Running                            
      
        Description
        -----------
        Gathers one running driver loaded on current system.
    #>
param([string]$ComputerName=$env:COMPUTERNAME)
Invoke-Command -ComputerName $ComputerName -ScriptBlock {
foreach ($driver in (gwmi -Query "Select Name,State,Description,DisplayName,PathName,ServiceType,StartMode FROM win32_systemdriver WHERE State='Running'"|select Name,State,Description,DisplayName,PathName,ServiceType,StartMode))
            {
                 
                $Driver.Pathname = ($driver.pathname.Replace("\??\", ""))                
                $Path = Get-ChildItem -Path $Driver.Pathname -ErrorAction SilentlyContinue|Select FullName,LastWriteTime,@{ Label = 'FileVersion'; Expression = { $_.VersionInfo.FileVersion }}
                                                           
             
                $Obj= New-Object -Type PSObject -Property @{
                      Name           = $Driver.Description
                      Path           = $Driver.Pathname
                      Date           = $Path.LastWriteTime
                      Version        = $Path.FileVersion
                      State          = $Driver.State
                      Type           = $Driver.ServiceType
                                                       }
      $Obj|Where {$_.Version -notlike "*(win*"}    
            }
                                                            }
    }
    
function Get-NetFrameworkVersion {
[CmdletBinding()]
param($ComputerName = $env:COMPUTERNAME)
$dotNetRegistry  = 'SOFTWARE\Microsoft\NET Framework Setup\NDP'
$dotNet4Registry = 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
$dotNet4Builds = @{
       30319  = '.NET Framework 4.0'
       378389 = '.NET Framework 4.5'
       378675 = '.NET Framework 4.5.1 (8.1/2012R2)'
       378758 = '.NET Framework 4.5.1 (8/7 SP1/Vista SP2)'
       379893 = '.NET Framework 4.5.2'
       380042 = '.NET Framework 4.5 and later with KB3168275 rollup'
       393295 = '.NET Framework 4.6 (Windows 10)'
       393297 = '.NET Framework 4.6 (NON Windows 10)'
       394254 = '.NET Framework 4.6.1 (Windows 10)'
       394271 = '.NET Framework 4.6.1 (NON Windows 10)'
       394802 = '.NET Framework 4.6.2 (Windows 10 Anniversary Update)'
       394806 = '.NET Framework 4.6.2 (NON Windows 10)'
       460798 = '.NET Framework 4.7 (Windows 10 Creators Update)'
       460805 = '.NET Framework 4.7 (NON Windows 10)'
}
foreach($Computer in $ComputerName) {
try
    {
        Test-Connection -ComputerName $Computer -Count 1 -ErrorAction Stop|Out-Null
        $builtInWinFeature = Get-WmiObject -ComputerName $Computer -query "SELECT Caption,Version FROM Win32_OperatingSystem" -ErrorAction Stop      
    }
  
catch
    {
    #Write-Warning ("$($Computer)- WMI: " + $_.Exception.Message)
    #$builtInWinFeature = $null
    }  
  
    #Check if computer is running 2012 or higher. Checking for built-in .NET 4.x installs, since they don't register as the other installable versions.  
    if (($builtInWinFeature.Caption -like "*server*") -and ($builtInWinFeature.Caption -notlike "*200*")){      
        $netCollection = @()
  
        #handle errors if Get-WinFeature is not found on machine running script, but $computer is a Server 2012 or greater system
        try {
                $serverDotNet = Get-WindowsFeature -ComputerName $Computer -Name NET* -ErrorAction Stop|Select *|Where {$_.Installed -eq $true -and $_.Parent -eq $null} |Select DisplayName,Installed
                }
        catch {
                #"Error loading Get-WindowsFeature. Server $($Computer) integrated .NET installs will not be shown"
                #$serverDotNet = $null
                }
                        foreach ($feature in $serverDotNet) {
                        $obj = New-Object -Type PSObject -Property @{
                               ComputerName = $Computer
                               NetFXVersion = $feature.Displayname
                               NetFxBuild   = ($builtInWinFeature.Caption -replace "Microsoft ","") + " (Integrated)"
                               }
                        }
                  
                        #Only write to netCollection if Get-WindowsFeature worked
                        if ($serverDotNet) {$netCollection += $obj}
  
        #loop through registry
        if($regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)) {
              if ($netRegKey = $regKey.OpenSubKey("$dotNetRegistry")) {
                     foreach ($versionKeyName in $netRegKey.GetSubKeyNames()) {
                           if ($versionKeyName -match '^v[123]') {
                                  $versionKey = $netRegKey.OpenSubKey($versionKeyName)
                                  $version = Invoke-Command -scriptblock { $erroractionpreference = "SilentlyContinue"
                                  [version]($versionKey.GetValue('Version', ''))}
                                  $objnet = New-Object -TypeName PSObject -Property @{
                                         ComputerName = $Computer
                                         NetFXBuild = $version.Build
                                         NetFXVersion = '.NET Framework ' + $version.Major + '.' + $version.Minor
                                  }
                    if ($objnet.NetFXBuild -ne $null){ #added to prevent null entries and errors from 2016 and greater O/S without .NET 3.x reg values
                        $netCollection += $objnet
                            } #end if netfxbuild $null check
                                        
                        }#endif versionkeyname foreach
                   } #end foreach $versionKeyName
              }#endIf $netRegKey
              if ($net4RegKey = $regKey.OpenSubKey("$dotNet4Registry")) {
                     if(-not ($net4Release = $net4RegKey.GetValue('Release'))) {
                           $net4Release = 30319
                     }
                     $objnet4 = New-Object -TypeName PSObject -Property @{
                           ComputerName = $Computer
                           NetFXBuild = $net4Release
                           NetFXVersion = $dotNet4Builds[$net4Release]
                     }
         
            $netCollection += $objnet4
                      
              }#end if $dotNet4RegKey
       }#end if $regkey
        $netCollection|select ComputerName, NetFXVersion, NetFXBuild
    } #end 2012 or greater 'if' loop
    else {
     
       try
            {
                Test-Connection -ComputerName $Computer -Count 1 -ErrorAction Stop| Out-Null
                if($regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)) {
              if ($netRegKey = $regKey.OpenSubKey("$dotNetRegistry")) {
                     foreach ($versionKeyName in $netRegKey.GetSubKeyNames()) {
                           if ($versionKeyName -match '^v[123]') {
                                  $versionKey = $netRegKey.OpenSubKey($versionKeyName)
                                  $version = [version]($versionKey.GetValue('Version', ''))
                                  New-Object -TypeName PSObject -Property @{
                                         ComputerName = $Computer
                                         NetFXBuild = $version.Build
                                         NetFXVersion = '.NET Framework ' + $version.Major + '.' + $version.Minor
                                  } | Select-Object ComputerName, NetFXVersion, NetFXBuild
                           }
                     }
              }
              if ($net4RegKey = $regKey.OpenSubKey("$dotNet4Registry")) {
                     if(-not ($net4Release = $net4RegKey.GetValue('Release'))) {
                           $net4Release = 30319
                     }
                     New-Object -TypeName PSObject -Property @{
                           ComputerName = $Computer
                           NetFXBuild = $net4Release
                           NetFXVersion = $dotNet4Builds[$net4Release]
                     } | Select-Object ComputerName, NetFXVersion, NetFXBuild
              }
       }
            }
        catch
            {
                #Write-Warning ("$($Computer)- Registry: " + $_.Exception.Message)
            }
  
    } #endElse
}#end foreach
}#end function    
    
#endregion Functions
       #region System Info
    Create-RAMPieChart -ComputerName $computer
    $RAMImgBits = [convert]::ToBase64String((Get-Content ($env:tmp + "\RAM-"+$computer + ".png") -Encoding Byte))
    $RAMImgHTML = "<img src=data:image/png;base64,$($RAMImgBits) alt='RAM' />"
    Remove-Item -Path ($env:tmp + "\RAM-"+$computer + ".png")
   
    Create-CPUPieChart -ComputerName $computer
    $CPUImgBits = [convert]::ToBase64String((Get-Content ($env:tmp + "\CPU-"+$computer + ".png") -Encoding Byte))
    $CPUImgHTML = "<img src=data:image/png;base64,$($CPUImgBits) alt='CPU' />"
    Remove-Item -Path ($env:tmp + "\CPU-"+$computer + ".png")
   
   
       $CompInfo = Get-WmiObject Win32_ComputerSystem -computername $computer
    $SysMan = $CompInfo.Manufacturer
   
     $SystemInfoTable = if (($SysMan -like "*Vmware*") -or ($SysMan -like "*Microsoft*") -or ($SysMan -like "*Xen")) {
                        Get-SystemInformation -ComputerName $computer|
                            ConvertTo-Html "System Manufacturer",
                                            "System Model",                                           
                                            "CPU Name",
                                            "Physical CPU Cores",
                                            "Logical CPU Cores",
                                            "System Uptime",
                                            "Operating System",
                                            "Power Plan",
                                            "CPU % Usage",
                                            "Total RAM (GB)",
                                            "Free RAM (GB)",
                                            "Free RAM %"`
                                            -As List -Fragment | ForEach {
                          $_ -replace "<table>","<table class=`"list`">"
                               }
                        }
               
                else {
               
                Get-SystemInformation -ComputerName $computer|
                            ConvertTo-Html "System Manufacturer",
                                            "System Model",
                                            "BIOS Vendor",
                                            "BIOS Version",
                                            "BIOS Date",
                                            "CPU Name",
                                            "Physical CPU Cores",
                                            "Logical CPU Cores",
                                            "System Uptime",
                                            "Operating System",
                                            "Power Plan",
                                            "CPU % Usage",
                                            "Total RAM (GB)",
                                            "Free RAM (GB)",
                                            "Free RAM %"`
                                             -As List -Fragment | ForEach {
                          $_ -replace "<table>","<table class=`"list`">"
                               }
               
                } 
   
   
    $GetMemoryStats = if (($SysMan -like "*Vmware*") -or ($SysMan -like "*Microsoft*") -or ($SysMan -like "*Xen"))
                         {
                            $Virtual="True"
                            Get-MemoryStats -ComputerName $Computer|select @{l="Virtual"; e={$Virtual}},Bank,"Capacity (GB)","Data Width","Form Factor","Memory Type"|
                            ConvertTo-Html -Fragment |  ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
                         }
                      
                       else {
                             Get-MemoryStats -ComputerName $Computer|select Bank,Manufacturer,"Capacity (GB)","Speed","Part Number","Serial Number","Data Width","Form Factor","Memory Type"|
                             ConvertTo-Html -Fragment | ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
                            }
                           
    #Disks
    $PhysicalDiskInfo = Get-WmiObject -ComputerName $computer -query "SELECT Model,DeviceID,Size,Status FROM Win32_DiskDrive" |
    Select-Object Model,DeviceID, @{n='Size (GB)';e={"{0:n2}" -f ($_.size/1gb)}}|
    ConvertTo-Html -Fragment| ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
   
    $LogicalDiskInfo = Get-WMIObject -ComputerName $computer -Query "Select SystemName,VolumeName,Name,Size,FreeSpace FROM Win32_LogicalDisk Where DriveType=3"|
    Select-Object SystemName, VolumeName, Name, @{n='Size (GB)';e={"{0:n2}" -f ($_.size/1gb)}}, @{n='FreeSpace (GB)';e={"{0:n2}" -f ($_.freespace/1gb)}}, @{n='PercentFree';e={"{0:n2}" -f ($_.freespace/$_.size*100)}} |
    ConvertTo-HTML -Fragment|  ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
       
    $PhysicalDiskPerf = Get-WmiObject -ComputerName $computer Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "NOT Name LIKE '_Total'" |
    Select @{Expression={$_.Name}; Label="Disk# Name" },AvgDiskQueueLength, @{Expression={$_.PercentIdleTime}; Label="% Idle Time"}, @{Expression={$_.PercentDiskReadTime}; Label="% Read Time"}, @{Expression={$_.PercentDiskWriteTime}; Label="% Write Time"}, @{Expression={$_.DiskWritesPersec * 1}; Label="Write IOPS"}, @{Expression={$_.DiskReadsPersec * 1}; Label="Read IOPS" }|
    ConvertTo-Html -Fragment|  ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
                                       
       #endregion System Info
       
    #region Proc&Services
       try {
    $owners = @{}
    gwmi -ComputerName $computer -Query "SELECT Handle FROM win32_process" |% {$owners[$_.handle] = $_.getowner().user}
    }
    catch {
        Write-Warning "Unable to retrieve process owners. Top 10 Process Report - Owners column may be inaccurate"
       }
       Start-Sleep -Milliseconds 50
   
    #create Process bar chart
    Create-ProcBarChart -ComputerName $computer -ProcessNumber $ProcessNumToFetch|Out-Null
    $ProcBarImgBits = [convert]::ToBase64String((Get-Content ($env:tmp + "\ProcBar-"+$computer + ".png") -Encoding Byte))
    $ProcBarImgHTML = "<img src=data:image/png;base64,$($ProcBarImgBits) alt='TopProcesses' />"
    Remove-Item -Path ($env:tmp + "\ProcBar-"+$computer + ".png")
   
    $TopProcesses = Get-WmiObject -ComputerName $computer -Query "SELECT Description,ProcessID,WorkingSetSize FROM Win32_Process"| Sort WorkingSetSize -Descending | Select @{l="Name"; e={$_.Description}},@{L="PID"; e={$_.ProcessID}},@{L="WS (MB)";E={[int64]($_.WorkingSetSize/1MB)}},@{label="Owner";e={$owners[$_.Processid.tostring()]}} -first $ProcessNumToFetch | ConvertTo-Html -Fragment
       
       $StoppedServicesReport = @()
       $StoppedServices = Get-WmiObject -ComputerName $computer -Query "SELECT Name,State,StartMode FROM Win32_Service WHERE StartMode='Auto' AND State='Stopped'"
       foreach ($StoppedService in $StoppedServices) {
              $row = New-Object -Type PSObject -Property @{
                     Name = $StoppedService.Name
                     Status = $StoppedService.State
                     "Start Mode" = $StoppedService.StartMode
              }
              
       $StoppedServicesReport += $row
       
       }
       
       $StoppedServicesReport = $StoppedServicesReport | ConvertTo-Html -Fragment|  ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
              
    $NsServicesReport = @()
    $NsServices = Get-WmiObject -ComputerName $computer -Query "Select StartName,Name,State,Startmode FROM Win32_Service WHERE NOT StartName LIKE '%NT Authority%' AND NOT StartName LIKE '%localsystem%'"
   
    foreach ($NsService in $NsServices) {
              $row = New-Object -Type PSObject -Property @{
                     "Service Account" = $NsService.StartName
            Name = $NsService.Name
                     Status = $NsService.State
                     StartMode = $NsService.StartMode
           
              }
              
       $NsServicesReport += $row
       
       }
   
    $NsServicesReport = $NsServicesReport |
    ConvertTo-Html "Service Account",Name,StartMode,Status -Fragment| ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
           try
{
    $Connect=Invoke-Command -ComputerName $computer { 1 } -ErrorAction Stop
   
    $NonWinServices = Invoke-Command -ComputerName $computer -ScriptBlock {Get-ItemProperty HKLM:\System\CurrentControlSet\Services\* |
    where {($_.DisplayName -notlike "*Microsoft*" -and $_.ImagePath -ne $null -and $_.ObjectName -ne $null -and $_.Description -ne $null -and $_.Description -notlike "*@*" -and $_.DisplayName -notlike "*@*" )}}|
    select @{E={$_.DisplayName}; L="Service" }, Description, @{L="Command"; E={$_.ImagePath}}, @{E={$_.ObjectName}; L="Account"}|
    Sort Service|
    ConvertTo-Html -Fragment | ForEach {
                          $_ -replace "<table>","<table class=`"services`">"
                          }
}
catch
{
$NonWinServices = '<p style="color:red;">No PS remote connectivity</p>'     
       
}   
    $EnvVariables = Get-WMIObject -ComputerName $computer -Query "SELECT VariableValue,Name,UserName,SystemVariable FROM Win32_Environment"|
                    select Name,@{e={$_.VariableValue};l="Variable"},@{e={$_.UserName.Trim("<",">")};l="User Account"}|
                    ConvertTo-Html -Fragment| ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
              
       #endregion Proc&Services
   
    #region Tasks
    $tasks = (Get-RemoteScheduledTasks -ComputerName $computer).Tasks|
                                   where {($_.State -ne 'Disabled') -and `
                                   ($_.Enabled) -and `
                                   ($_.NextRunTime -ne 'None') -and `
                                   (!$_.Hidden) -and `
                                   ($_.Author -ne 'Microsoft Corporation')}|
                                   Select Name, Author, Description,@{Label="Last Run"; Expression={$_.LastRunTime}},@{Label="Next Run"; expression={$_.NextRunTime}},@{Label="Last Result"; expression={$_.LastTaskDetails}}
    $tasks = $tasks |ConvertTo-Html -Fragment|  ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
    #endregion Tasks
   
    #region Startup
    $Startup = Get-Startup -ComputerName $computer |
    ConvertTo-Html -Fragment|  ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
    #endregion Startup
   
    #region Event Logs Report
       $SystemEventsReport = @()
       $SystemEvents = Get-winevent -ComputerName $computer -FilterXml $SysEvtFilter |Sort-Object timecreated -Descending
       foreach ($event in $SystemEvents) {
              $row = New-Object -Type PSObject -Property @{
                     TimeGenerated = $event.TimeCreated
                     EntryType = $event.LevelDisplayName
                     Source = $event.ProviderName
                     ID = $event.Id
            Message = $event.Message
              }
              $SystemEventsReport += $row
       }
                     
       $SystemEventsReport = $SystemEventsReport | ConvertTo-Html EntryType,ID,Message,Source,TimeGenerated -Fragment
       $SystemEventsReport = $SystemEventsReport|ForEach {
                          $_ -replace "<td>Warning</td>","<td style=`"padding-left:5px;padding-right:10px;white-space:nowrap;background-color:#FBD95B`">$YellowBangHTML Warning</td>"
                          }
   
    $SystemEventsReport = $SystemEventsReport|ForEach {
                          $_ -replace "<td>Error</td>", "<td style=`"padding-left:5px;padding-right:10px;white-space:nowrap;background-color:#FB7171`">$RedBangHTML Error</td>"
                          }
                         
       $ApplicationEventsReport = @()
       $ApplicationEvents = Get-winevent -ComputerName $computer -FilterXml $AppEvtFilter |Sort-Object timecreated -Descending
       foreach ($event in $ApplicationEvents) {
              $row = New-Object -Type PSObject -Property @{
                     TimeGenerated = $event.TimeCreated
                     EntryType = $event.LevelDisplayName
                     Source = $event.ProviderName
                     ID = $event.Id
            Message = $event.Message
              }
              $ApplicationEventsReport += $row
       }
       
       $ApplicationEventsReport = $ApplicationEventsReport | ConvertTo-Html EntryType,ID,Message,Source,TimeGenerated -Fragment
   
    $ApplicationEventsReport = $ApplicationEventsReport|ForEach {
                          $_ -replace "<td>Warning</td>","<td style=`"padding-left:5px;padding-right:10px;white-space:nowrap;background-color:#FBD95B`">$YellowBangHTML Warning</td>"
                          }
   
    $ApplicationEventsReport = $ApplicationEventsReport|ForEach {
                          $_ -replace "<td>Error</td>", "<td style=`"padding-left:5px;padding-right:10px;white-space:nowrap;background-color:#FB7171`">$RedBangHTML Error</td>"
                          }
   
       #endregion Event Logs Report
       
    #region network
    $NICS = Get-NicInfo -ComputerName $computer| Select "Hostname",
            "Name",
            "Network",
            "Link Speed",
            "Up Time",
            "MAC Address",
            "IP Address",
            "DHCP Server",
            "Subnet Mask",
            "Default Gateway",
            "DNS Suffix",
            "DNS Servers"|ConvertTo-Html -As List -Fragment|  ForEach {
                          $_ -replace "<table>","<table class=`"list`">"
                          }
   
    $Routes = (Get-RemoteRouteTable -computername $computer).Routes |select Destination, Mask, NextHop, Persistent, Metric, @{l="Int. Index"; e={$_.InterfaceIndex}}, Type|
    Sort-Object -Property Metric| ConvertTo-Html -Fragment |  ForEach {
                          $_ -replace "<table>","<table class=`"services`">"
                          }  
    #endregion network
   
   
#region updates/software   
  $NETinfo = Get-NetFrameworkVersion -ComputerName $computer
  
  $NETinfoHtml = $NETinfo |ConvertTo-Html -Fragment|ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }  
    
    try
{
   
   $Connect=Invoke-Command -ComputerName $computer { 1 } -ErrorAction Stop
  
   $64Prog = Invoke-Command -cn $computer -ScriptBlock {Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
        Get-ItemProperty |Where {$_.DisplayName -ne $null}|
        Sort-Object -Property DisplayName |
        Select-Object -Property DisplayName, DisplayVersion, Publisher, InstallDate
        }
    $32Prog = Invoke-Command -cn $computer -ScriptBlock {Get-ChildItem -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
        Get-ItemProperty |Where {$_.DisplayName -ne $null}|
        Sort-Object -Property DisplayName |
        Select-Object -Property DisplayName, DisplayVersion, Publisher, InstallDate
        }
    $SoftwareReport = @()  
    $Software = $64Prog + $32Prog
        foreach ($Soft in $Software) {
            $row = New-Object -Type PSObject -Property @{
            Name = $Soft.DisplayName
            Publisher = $Soft.Publisher
                     Version = $Soft.DisplayVersion
                     InstallDate = $Soft.InstallDate
            }
        $SoftwareReport += $row
        }
    $InstalledSoftware = $SoftwareReport|Where {$_.Name -ne $null}| Sort-Object Name|
    Select-Object Name, Version, Publisher, @{Label="Installed On"; Expression={[DateTime]::ParseExact($_.InstallDate,'yyyyMMdd',$null)}}|
    ConvertTo-Html -Fragment|  ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
   
    $Drivers = Get-WindowsDriver -ComputerName $computer|Select Name,Version,Date,Path,Type,State|Sort Name |ConvertTo-Html -Fragment| ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
}
catch
{
       $InstalledSoftware = '<p style="color:red;">No PS remote connectivity</p>'
       $Drivers           = '<p style="color:red;">No PS remote connectivity</p>'
}   
 
    $Updates = Get-MSHotfix -ComputerName $computer|Sort-Object InstalledOn -descending
    $HTMLUpdates = $Updates|Select @{L="KB ID"; E={$_.HotFixID}}, @{n='KB Link';e={"<a href='$($_.Caption)'>$($_.Caption)</a>"}}, Description, @{L="Install Date"; E={$_.InstalledOn}}, @{L="Installed By User"; E={$_.InstalledBy}}|
    Sort-Object {$_."Install Date" -as [datetime]} -Descending|
    ConvertTo-Html -Fragment| ForEach {
                          $_ -replace "<table>","<table class=`"disks`">"
                          }
    Add-Type -AssemblyName System.Web
    [System.Web.HttpUtility]::HtmlDecode($HTMLUpdates)|Set-Variable -Name HTMLClickUpdates
   
    #endregion updates/software
   
   
       # Create HTML Report for the current System being looped through
       $CurrentSystemHTML = @"
       
       
    <h1>$computer Report</h1>
    <p><i>Report generated by Systems Report Tool version $version on $date.</i></p>
       <hr size=1 width="40%">
    <div style="width:50%; margin-left: auto; margin-right: auto;">
    <div class="square" style="float: left; width: 50%;">
    $SystemInfoTable
    </div>
    <div class="square" style="float: right; width:50%;">
       $CPUImgHTML
    $RAMImgHTML
       </div>
    </div>
   
   
   
    <input class="toggle-box" id="identifier-$computer-1" type="checkbox" name="grouped"><label for="identifier-$computer-1"> Disk & Memory Information- $computer </label> 
       <div>
    <h3>Memory Module Information</h3>
    $GetMemoryStats
    <br>
    <br>
    <h3>Disk Information and Performance</h3>
    <p>Physical Disks</p>
    $PhysicalDiskInfo
    <p>Physical Disk Performance Data</p>
    $PhysicalDiskPerf
       <p>Logical Disks</p>
       $LogicalDiskInfo
    <br>
    <br>
   
    </div>
          
          
    <input class="toggle-box" id="identifier-$computer-2" type="checkbox" name="grouped"><label for="identifier-$computer-2"> System Processes and Services - $computer</label> 
       <div>
   
    <h3>System Processes - Top $ProcessNumToFetch Highest Memory Usage</h3>
    <p>The following $ProcessNumToFetch processes are consuming the highest Working Set Memory* on $computer</p>
    <div style="width:50%; margin-left: auto; margin-right: auto;">
   
   
    <div class="square" style="float: left; width:50%;">
       
    $TopProcesses
    <p class="small"><font size=1><i>*Number reflects Working Set KB (see Resource Monitor) converted to MB</i></font></p>
       </div>  
   
    <div class="square" style="float: right; width:50%;">
       $ProcBarImgHTML
    </div>
    </div>
   
    <h3>System Services - Stopped</h3>
       <p>The following services are set to Automatic startup, yet are currently stopped on $computer</p>
       $StoppedServicesReport
    <br>
    <br>
   
    <h3>Startup Entries</h3>
    <p>Startup Entries registered on $computer</p>
    $Startup
    <br>
    <br>
   
    <h3>Non-Windows Services</h3>
    <p>The following services on $computer do not come standard with Windows</p>
    $NonWinServices
    <br>
    <br>
    <h3>Non-Standard Service Accounts</h3>
       <p>The following service accounts on $computer do not come standard with Windows</p>
       $NsServicesReport
       <br>
    <br>
       
    <h3>Scheduled Tasks Added to Windows</h3>
    <p>The following scheduled tasks on $computer do not come standard with Windows.</p>
    $tasks   
    <br>
    <br>
   
    <h3>Environment Variables</h3>
    <p>Environment Variables from $computer</p>
    $EnvVariables
    <br>
    <br>
   
       </div>
   
       
    <input class="toggle-box" id="identifier-$computer-3" type="checkbox" name="grouped"><label for="identifier-$computer-3"> Networking Information - $computer</label> 
       <div>
   
    <h3>Network Adapter Information</h3>
    <p>Configuration information of all active network adapters on $computer</p>
    $NICS
    <br> 
    <br>
       
    <h3>Route Tables</h3>
    <p>Routing table from $computer</p>
    $Routes
    <br>
    <br>
   
    </div>
   
   
    <input class="toggle-box" id="identifier-$computer-4" type="checkbox" name="grouped"><label for="identifier-$computer-4"> Event Log Report - $computer</label> 
       <div>
   
       <div class="square" style="float: left; width: 49%">
    <h3>System Log</h3>
    <p>List of System log events that had an Event Type of either Warning, Error, or Critical on $computer within 24 hours</p>
       $SystemEventsReport
    </div>
   
    <div class="square" style="float: right; width: 49%">
    <h3>Application Log</h3>
    <p>List of Application log events that had an Event Type of either Warning or Error on $computer within 24 hours</p>
       $ApplicationEventsReport
    </div>
   
    <br>
    <br>
    </div>
 
     
    <input class="toggle-box" id="identifier-$computer-5" type="checkbox" name="grouped"><label for="identifier-$computer-5"> Software Information - $computer</label> 
       <div>
   
   
   
    <input class="toggle-boxin" id="identifier-$computer-6" type="checkbox" name="grouped"><label for="identifier-$computer-6"> Installed Programs - $computer</label> 
              <div>
   
   <p>List .NET Framework Installations on $computer</p>
    $NETinfoHtml
   
    <p>List of installed software on $computer.</p>
    $InstalledSoftware
   
              </div>
   
   
    <input class="toggle-boxin" id="identifier-$computer-7" type="checkbox" name="grouped"><label for="identifier-$computer-7"> Installed Windows Updates - $computer</label>
              <div>
   
    <p>This lists all the Windows <i>(only)</i> Updates installed on $computer.<p>
    $HTMLClickUpdates
   
              </div>
   
    <input class="toggle-boxin" id="identifier-$computer-8" type="checkbox" name="grouped"><label for="identifier-$computer-8"> Loaded Windows Drivers - $computer</label>
              <div>
   
    <p>This lists all third-party kernel drivers currently running on $computer.<p>
    $Drivers
   
              </div>
   
    <br>
    <br>
     
   
    </div> 
 <br>
 <br>
"@
$CurrentSystemHTML
}
$computers | % {Start-Job -Scriptblock $scriptblock -ArgumentList $_ | Out-Null}
""
Write-Host -ForegroundColor Yellow "Waiting for jobs to finish... This may take awhile, depending on amount of servers queried."
$HTMLMiddle = Get-Job | Wait-Job | Receive-Job
""
Write-Host -ForegroundColor Yellow "Gathering job results..."
#region Assemble the HTML Header and CSS for Report
$HTMLHeader = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>NextGen System Health Report</title>
<style type="text/css">
<!--
body
       {
              font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
              text-align: center;
       }
 
table
       {
        margin-left:auto;
              margin-right:auto;
        border: 1px solid rgb(190, 190, 190);
        border-radius: 2px;
              border-spacing: 0px;
        font-Family: Tahoma, Helvetica, Arial;
        font-Size: 8pt;
        text-align: left;
    }
th
    {
        Text-Align: Left;
              background-color: #F3F3F3;
        Padding: 4px;
              font-weight: normal;
              border-bottom: 1px solid rgb(190, 190, 190);
    }
   
tr:hover td
    {
        background-color: DodgerBlue ;
        Color: #F5FFFA;
           }
tr:nth-child(odd)
       {
              background-color:#F9F9F9;
       }      
td
    {
        Vertical-Align: middle;
        Padding: 4px;
    }
h1
       {
              clear: both;
              font-size: 30px;
              font-weight: 300;
              font-family: Verdana, Helvetica, Arial;
       }
h2
       {
              clear: both;  
              font-size: 22px;
              font-weight: 300;
              margin-bottom: 10px;
              margin-top: 30px;
              background-color: #E5E5E5;
              width: 75%;
              margin: auto;
              margin-bottom: 15px;
              border:2px solid #CCCCCC;
              padding: 10px;
       }
h3
       {
              clear: both;
              border-radius: 10px;
              font-family: Verdana, Arial;
              font-size: 17px;
              font-weight: 300;
              margin-bottom: 10px;
              margin-top: 10px;
              width: 40%;
              margin: auto;
              margin-bottom: 10px;   
              padding: 10px;
       }
p
       {
              margin-top: 10px;
              margin-left: 0px;
              font-size: 12px;
              font-weight: 300;    
              text-align: center;
              margin-bottom: 5px;
       }
p.small
       {
              width: 175px;
              margin: auto;
              margin-top: 0px;
              font-size: 12px;
              text-align: center;
              margin-bottom: 10px;
    }
       
table.disks
       {
              table-layout:fixed;
              min-width:50%;
              max-width: 50%;
       }
table.services
       {
              table-layout:fixed;
              min-width:60%;
              max-width: 60%;
       }
table.list
       {
              float: center;
       }
table.list tr:nth-child(odd)
       {
              background-color:#F9F9F9;
       }
table.list td:nth-child(1)
       {
              font-weight: 600;
              border-right: 1px solid rgb(190, 190, 190);
              text-align: right;
       }
.square
       {
              width:100%;
              height:auto;
              margin-bottom: 15px;
              margin-top: 15px;
              text-align:center;
       }
.toggle-box
       {
              display: none;
       }
.toggle-box + label
       {
              cursor: pointer;
              display: block;
              clear: both;  
              font-size: 18px;
              font-weight: 300;
              margin-bottom: 15px;
              margin-top: 15px;
              background-color: #f3f3f3;
              width: 50%;
              margin-left: 25%;
              text-align:left;
              border-radius: 5px;
              box-shadow: 1px 1px #CCCCCC;
              padding: 10px;
       }
.toggle-box + label + div
       {
              display: none;
              margin-bottom: 10px;
              margin-left:auto;
              margin-right:auto;
       }
.toggle-box:checked + label + div
       {
              display: block;
              margin-left:18px;
              margin-right:auto;
       }
.toggle-box + label:before
       {
              content: "+";
              display: block;
              float: left;
              font-weight: 300;
              font-size:17px;
              line-height: auto;
              margin-right: 5px;
              margin-left:auto;
              text-align: center;
              width: 23px;
              height: 23px;
       }
.toggle-box:checked + label:before
       {
              content: "\2212";
       }
.toggle-boxin
       {
              display: none;
       }
.toggle-boxin + label
       {
              cursor: pointer;
              display: block;
              clear:both;
              font-size:16px;
              font-weight: 300;
              margin-bottom: 10px;
              margin-top: 10px;
              background-color:#F9F9F9;
              width: 35%;
              margin-left: 30%;
              text-align:left;
              border-radius: 5px;
              box-shadow: 1px 1px #CCCCCC;
              padding: 10px;
       }
.toggle-boxin + label + div
       {
              display: none;
              margin-bottom: 10px;
              margin-left:auto;
              margin-right:auto;
       }
.toggle-boxin:checked + label + div
       {
              display: block;
       }
.toggle-boxin + label:before
       {
              content: "+";
              display: block;
              float: left;
              font-weight: 300;
              font-size:15px;
              line-height: auto;
              margin-right: 5px;
              margin-left:auto;
              text-align: center;
              width: 21px;
              height: 21px;
       }
.toggle-boxin:checked + label:before
       {
              content: "\2212";
       }
-->
</style>
</head>
<body>
"@
#endregion Assemble the HTML Header and CSS for Report
# Assemble the closing HTML for our report.
$HTMLEnd = @"
</body>
</html>
"@
# Assemble the final report from all our HTML sections
""
Write-Host -ForegroundColor Yellow "Assembling HTML for final report..."
$HTMLmessage = $HTMLHeader + $HTMLMiddle + $HTMLEnd
# Save the report out to a file in the current path
Add-Type -AssemblyName System.Web
    [System.Web.HttpUtility]::HtmlDecode($HTMLmessage) | Out-File ((Get-Location).Path + "\NGSystemReport.html")
""
Write-Host -ForegroundColor Green "Complete!"
""
""
Write-Host -ForegroundColor DarkYellow "Please zip the contents of this folder:" (Get-Location).Path "(without the cmd and ps1 file), for transfer."
# Email our report out
# send-mailmessage -from $fromemail -to $users -subject "Systems Report" -Attachments $ListOfAttachments -BodyAsHTML -body $HTMLmessage -priority Normal -smtpServer $server
 