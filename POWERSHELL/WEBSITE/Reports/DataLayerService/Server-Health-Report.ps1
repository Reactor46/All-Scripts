##Windows Server Health Report (Date: 30.Jun.2017)																							
##This script is created for the Windows server health report status for windows all flavors.										
##The first Author was: Prashant Dev Pandey Email ID: prashantdev.pandey@gmail.com.													
##Second Author: Manan.Shastri Email ID: manan.shastri@outlook.com. 																
##I have Modified it to generate more report Event Log, Task Scheduler status and current CPU Load and send an email in HTML format.  
##These scripts provide a report in HTML format of the server average CPU and Memory utilization along with Disk space utilization,	
##Paging Size, Performance counter, Event Log and  Task Scheduler status.															
##You have to execute file Server-Health-Report.ps1 and it will give HTML report. You can setup task scheduler for a daily report.	
##In this script, you have to change destination path of HTML file where you want to generate a report.  							

$Outputreport="Test"
#Declaring CSS
$Outputreport +="<style>TABLE{ border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;align:center;width:100%;}
TH{border-width: 1px;background-color: lightblue;bgcolor=blue;padding: 3px;border-style: solid;border-color: black;}
TD{border-width: 1px;color: white;background-color: gray;padding: 3px;border-style: solid;border-color: black;}
 
h1{text-shadow: 1px 1px 1px #000,3px 3px 5px blue; text-align: center;font-style: calibri;font-family: Calibri;</style>"



## Get Host Name
$Hostname = Test | Out-String
## Get version
$Version = (Get-WmiObject -class Win32_OperatingSystem).Caption | Out-String

## Get Uptime
$UPTIME=Get-WmiObject Win32_OperatingSystem
$up = [Management.ManagementDateTimeConverter]::ToDateTime($UPTIME.LastBootUpTime) | Out-String

## Get Disk Spaces
$Disk = Get-WmiObject Win32_logicaldisk -ComputerName LocalHost -Filter "DriveType=3" |select -property DeviceID,@{Name="Size(GB)";Expression={[decimal]("{0:N0}" -f($_.size/1gb))}},@{Name="Free Space(GB)";Expression={[decimal]("{0:N0}" -f($_.freespace/1gb))}},@{Name="Free (%)";Expression={"{0,6:P0}" -f(($_.freespace/1gb) / ($_.size/1gb))}}|ConvertTo-Html

## Get Critical Service Status here i have given SQL service you can pass different service name as per your requirement
$Private:wmiService =gsv -include "*SQL*" -Exclude "*ySQL*","*spo*"|select Name,DisplayName,Status|ConvertTo-Html
#$Services =gsv -include "*SQL*","*MpsSvc*" -Exclude "*ySQL*","*spo*"|select Name,DisplayName,Status|ConvertTo-Html
$Services =gsv -include "*SQL*","*FileZilla Server*","*MpsSvc*" -Exclude "*ySQL*","*spo*"|select Name,DisplayName,Status|ConvertTo-Html

## Get CPU Utilization
$CPU_Utilization = Get-Process|Sort-object -Property CPU -Descending| Select -first 15 -Property ID,ProcessName,@{Name = 'CPU In (%)';Expression = {$TotalSec = (New-TimeSpan -Start $_.StartTime).TotalSeconds;[Math]::Round( ($_.CPU * 100 /$TotalSec),2)}},@{Expression={$_.threads.count};Label="Threads";},@{Name="Mem Usage(MB)";Expression={[math]::round($_.ws / 1mb)}},@{Name="VM(MB)";Expression={"{0:N3}" -f($_.VM/1mb)}}|ConvertTo-Html

 

## Get Each Processor Utilization
$arr=@()
$ProcessorObject=gwmi win32_processor
foreach($processor in $ProcessorObject)
{
   $arr += $processor.Caption
   $arr += $processor.LoadPercentage
}

## Security Patches
$SecPatch = get-hotfix -Description "Security Update" |sort "Description" -desc | select Description,installedon -first 1 | Out-String

## RAM Usage
$Private:perfmem = Get-WmiObject -namespace root\cimv2 Win32_PerfFormattedData_PerfOS_Memory
$Private:totmem = Get-WmiObject -namespace root\cimv2 CIM_PhysicalMemory 
[Int32]$Private:totalcapacity = 0 
foreach ($Mem in $totmem) 
{ 
$totalcapacity += $Mem.Capacity / 1Mb 
} 
#Get-WmiObject Win32_PhysicalMemory | ForEach-Object {$totalcapacity += $_.Capacity / 1Mb} 

$Private:tmp = New-Object -TypeName System.Object 
$tmp | Add-Member -Name CapacityMB -Value $totalcapacity -MemberType NoteProperty 
$tmp | Add-Member -Name AvailableMB -Value $perfmem.AvailableMBytes -MemberType NoteProperty
$ram_usage = $tmp |ConvertTo-Html

## Physical Memory
function Get-MemoryUsage ($ComputerName=$ENV:ComputerName) {
if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
$ComputerSystem = Get-WmiObject -ComputerName $ComputerName -Class Win32_operatingsystem -Property TotalVisibleMemorySize, FreePhysicalMemory
$MachineName = $ComputerSystem.CSName
$FreePhysicalMemory = ($ComputerSystem.FreePhysicalMemory) / (1mb)
$TotalVisibleMemorySize = ($ComputerSystem.TotalVisibleMemorySize) / (1mb)
$TotalVisibleMemorySizeR = "{0:N2}" -f $TotalVisibleMemorySize
$TotalFreeMemPerc = ($FreePhysicalMemory/$TotalVisibleMemorySize)*100
$TotalFreeMemPercR = "{0:N2}" -f $TotalFreeMemPerc
# print the machine details:
"<table border=1 width=100>"
"<tr><th>RAM</th><td>$TotalVisibleMemorySizeR GB</td></tr>"
"<tr><th>Free Physical Memory</th><td>$TotalFreeMemPercR %</td></tr>"
"</table>"

}}
$PhyMem = Get-MemoryUsage
$Hotfix = (Get-HotFix | Sort-Object InstalledOn -ErrorAction SilentlyContinue)| Select-Object -last 50 HotFixID,InstalledBy,InstalledOn,Description | ConvertTo-Html
$Processor_Counter = Get-Counter "\Processor(_total)\% Processor Time" | ConvertTo-Html
$Total_Threads = (Get-Process |Select-Object -ExpandProperty Threads).Count
## Paging

function Get-PageFile { 
param( 
    [string]$computer="." 
)    
        Get-WmiObject -Class Win32_PageFileUsage  -ComputerName $computer | 
        Select  @{Name="File";Expression={ $_.Name }}, 
        @{Name="Base Size(MB)"; Expression={$_.AllocatedBaseSize}}, 
        @{Name="Peak Size(MB)"; Expression={$_.PeakUsage}},  
        TempPageFile 
  }
  
$PhysicalRAM = (Get-WMIObject -class Win32_PhysicalMemory  |
Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)})
$ht = @{}
$ht.Add('Total_Ram(GB)',$PhysicalRAM)
$OSRAM = gwmi Win32_OperatingSystem  |
foreach {$_.TotalVisibleMemorySize,$_.FreePhysicalMemory}
$ht.Add('Total Visable RAM GB',([Math]::Round(($OSRAM[0] /1MB),4)))
$ht.Add('Available_Ram(GB)',([Math]::Round(($OSRAM[1] /1MB),4)))
$RAM = New-Object -TypeName PSObject -Property $ht|ConvertTo-Html
$Paging1=Get-PageFile|ConvertTo-Html
$Paging =  Get-WMIObject Win32_PageFileSetting |  select Name, InitialSize, MaximumSize|ConvertTo-Html

$Available_Bytes=Get-Counter -Counter "\Memory\Available Bytes"|Select -ExpandProperty CounterSamples|Select CookedValue |ft -HideTableHeaders|out-string
$att=Get-Counter -Counter "\Memory\Committed Bytes"|Select -ExpandProperty CounterSamples|Select CookedValue |ft -HideTableHeaders|out-string
$Comitted_Bytes="{0:N0}" -f (($att.trim())/1024/1024)
$Handle_Count=Get-Counter -Counter "\Process(_Total)\Handle Count"|Select -ExpandProperty CounterSamples|Select CookedValue |ft -HideTableHeaders|out-string
$Thread_Count=Get-Counter -Counter "\Process(_Total)\Thread Count"|Select -ExpandProperty CounterSamples|Select CookedValue |ft -HideTableHeaders|out-string
$ftt=Get-Counter -Counter "\memory\Pool Paged Bytes"|Select -ExpandProperty CounterSamples|Select CookedValue |ft -HideTableHeaders|out-string
$Pool_Paged="{0:N0}" -f (($ftt.trim())/1024/1024)
$dtt=Get-Counter -Counter "\memory\Pool Nonpaged Bytes"|Select -ExpandProperty CounterSamples|Select CookedValue |ft -HideTableHeaders|out-string
$Pool_NonPaged="{0:N0}" -f (($dtt.trim())/1024/1024)
$Total_process=(get-process).count
$d=get-date

$events = Get-WmiObject -Class Win32_NTLogEvent -filter 'Type="error" or Type="warning" and (logfile="system" or logfile="application")'|Select-Object -first 60 -Property  Type, Logfile, EventCode, Message, TimeGenerated |ConvertTo-Html
$CPULoad = Get-WmiObject  win32_processor |  Measure-Object -property LoadPercentage -Average | Select Average |ConvertTo-Html

$sched = New-Object -Com "Schedule.Service"
$sched.Connect()
$out = @()
$sched.GetFolder("\").GetTasks(0) | % {
    $xml = [xml]$_.xml
    $out += New-Object psobject -Property @{
        "Name" = $_.Name
        "Status" = switch($_.State) {0 {"Unknown"} 1 {"Disabled"} 2 {"Queued"} 3 {"Ready"} 4 {"Running"}}
        "NextRunTime" = $_.NextRunTime
        "LastRunTime" = $_.LastRunTime
        "LastRunResult" = $_.LastTaskResult
        "Author" = $xml.Task.Principals.Principal.UserId
        "Created" = $xml.Task.RegistrationInfo.Date
    }
}

$out | Select-Object  Name, Status, NextRuNTime, LastRunTime, LastRunResult, Author, Created | ConvertTo-Html >C:\inetpub\wwwroot\task-sceduler.html






#This is the HTML view you can customize accoring to your requirement .
$Outputreport +="<BODY><HTML>"
$Outputreport +="<h2 align=center><u>SERVER HEALTH CHECK REPORT AS ON $d</u></h2>"
$Outputreport +="</br></br>"
$Outputreport +="<table border=1 ><tr><td>"
$Outputreport +="<table border=1 width=100%>"
$Outputreport +="<tr><th><B>Hostname</B></th><td>"+$Hostname+"</td></tr>"
$Outputreport +="<tr><th><B>Version</B></th><td>"+$Version+"</td></tr>"
$Outputreport +="<tr><th><B>Uptime</B></th><td>"+$up+"</td></tr>"
$Outputreport +="<tr><th><B>Physical Memory(MB)</B></th><td>"+$RAM+"</td></tr></td></tr><tr><td><tr><th><B>System</B></th></tr><tr><th>Total Handles</th><td>"+$Handle_Count.trim()+"</td></tr><tr><th>Total Thread</th><td>"+$Thread_Count.trim()+"</td></tr><tr><th>Total Process</th><td>"+$Total_process+"</td></tr><tr><th>Commit(MB)</th><td>"+$Comitted_Bytes.trim()+"</td></tr></td>"
$Outputreport +="<td><tr><th><B>Kernel Memory(MB)</B></th></tr><tr><th>Paged</th><td>"+$Pool_Paged.trim()+"</td></tr><tr><th>Non Paged</th><td>"+$Pool_NonPaged.trim()+"</td></tr></td></tr></table>"
$Outputreport += "</table></BODY></HTML>"
$Outputreport +="</br>"
$Outputreport +="</br>"
$Outputreport +="<BODY><HTML>"
$Outputreport +="<table border=1 width=50%>"
$Outputreport +="<tr><th><B>Disk Size</B></th><td>"+$Disk+"</td></tr>"
$Outputreport +="<tr><th><B>Services</B></th><td>"+$Services+"</td></tr>"
$Outputreport +="<tr><th><B>Top15 Process</B></th><td>"+$CPU_Utilization+"</td></tr>"
$Outputreport +="<tr><th><B>Ram_Usage</B></th><td>"+$ram_usage+"</td></tr>"
$Outputreport +="<tr><th><B>Physical Memory</B></th><td>"+$PhyMem+"</td></tr>"
$Outputreport +="<tr><th><B>Processor Counter</B></th><td>"+$Processor_Counter+"</td></tr>"
$Outputreport +="<tr><th><B>Last 15 HotFix</B></th><td>"+$Hotfix+"</td></tr>"
$Outputreport +="<tr><th><B>Paging</B></th><td>"+$Paging+"</td></tr>"
$Outputreport +="<tr><th><B>No Of Threads</B></th><td>"+$Total_Threads+"</td></tr>"
$Outputreport +="<tr><th><B>Average process</B></th><table border=1><tr><th>Cpu Load%</th></th></tr><tr><td>"+$CPULoad+"</td></tr></table></tr>"
$Outputreport += "</table></BODY></HTML>"
$Outputreport +="<tr><th><B>100Events</B></th><td>"+$events+"</td></tr>"
$Outputreport +="<BODY><HTML>"
$Outputreport +="<tr><th><B>Taskscheduler</B></th></tr>"
$Outputreport += "</table></BODY></HTML>"
$Outputreport | out-file C:\inetpub\wwwroot\Windows_Server_Health_Status.html
Invoke-Expression C:\inetpub\wwwroot\Windows_Server_Health_Status.html
<##Send email functionality from below line, use it if you want   
$smtpServer = ""	#Add SMTP server
$smtpFrom = ""  #Add From ID
$smtpTo =  'abc@think.tank, xyz@think.tank' #Add your email id
$messageSubject = "TEST-Servers Health report"
$message = New-Object System.Net.Mail.MailMessage $smtpFrom, $smtpTo
$message.Subject = $messageSubject
$message.IsBodyHTML = $true
$message.Body = "<head><pre>$style</pre></head>"
$message.Body += Get-Content # Path Of html file like D:\ps\Scripts\Windows_Server_Health_Status.html
$message.Body += Get-Content # Path Of html file like D:\ps\Scripts\task-sceduler.html
$smtp = New-Object Net.Mail.SmtpClient($smtpServer, 587)
$smtp.Send($message)
"Email Sent"#>