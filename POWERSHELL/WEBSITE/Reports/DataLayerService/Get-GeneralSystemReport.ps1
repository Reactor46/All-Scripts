##############################################################################
# NAME: Get-GeneralSystemReport.ps1
#
# AUTHOR(s): Sean Duffy (mostly)
#            Jeffery Hicks (a good part)
#            Zachary Loeber (just a bit of hacking together of things)
# DATE:   09/19/2012
# EMAIL:  zloeber@gmail.com
#
# COMMENT:  
#   Based on Sean's daily system report found here:
#   http://www.simple-talk.com/sysadmin/powershell/building-a-daily-systems-report-email-with-powershell/
#   I removed dependancies on the chart requirements and replaced graphs with
#   html code from Jeffery Hicks code found here:
#   http://jdhitsolutions.com/blog/2012/02/create-html-bar-charts-from-powershell/
#	I also added parameters to make the entire script more portable
# VERSION HISTORY
# 1.1: 11/12/2012
#     - Added some parameters
#     - Changed some default parameters
# 1.0: 09/19/2012 Initial Version.
# 
# TO ADD
#   - ???
##############################################################################

#region Parameters
[CmdletBinding()]
param
(
	[Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[AllowEmptyString()]
	[Alias('Server Name')]
	[String]$ServerName=".",				
	[Parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
	[Alias('Email Relay')]
	[String]$EmailRelay = "localhost",		
	[Parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
	[Alias('Email Sender')]
	[String]$EmailSender='systemreport@localhost',
	[Parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
	[Alias('Email Recipient')]
	[String]$EmailRecipient='default@yourdomain.com',
	[Parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
	[Alias('Send Mail')]
	[Bool]$SendMail=$false,
	[Parameter(Position=5,Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
	[Bool]$SaveReport=$true,
	[Parameter(Position=6,Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
	[String]$ReportName=".\GeneralSystemStatus.html"
)
#endregion Parameters

#region Configuration
## Environment Specific - Change These ##
$EventNum = 0         # Number of events to fetch for system report
$ProccessNumToFetch = 10   # Number of processes to fetch for system report

## Required - Leave These Alone ##
# System and Error Report Headers
$HTMLHeader = @'
<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Frameset//EN' 'http://www.w3.org/TR/html4/frameset.dtd'>
<html><head><title>My Systems Report</title>
<style type='text/css'>
<!--
body {
font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}

    #report { width: 835px; }

    table{
   border-collapse: collapse;
   border: none;
   font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
   color: black;
   margin-bottom: 10px;
}
   table td{
   font-size: 12px;
   padding-left: 0px;
   padding-right: 20px;
   text-align: left;
}
   table th {
   font-size: 12px;
   font-weight: bold;
   padding-left: 0px;
   padding-right: 20px;
   text-align: left;
}

h2{ clear: both; font-size: 130%; }

h3{
   clear: both;
   font-size: 115%;
   margin-left: 20px;
   margin-top: 30px;
}

p{ margin-left: 20px; font-size: 12px; }

table.list{ float: left; }
   table.list td:nth-child(1){
   font-weight: bold;
   border-right: 1px grey solid;
   text-align: right;
}

table.list td:nth-child(2){ padding-left: 7px; }
table tr:nth-child(even) td:nth-child(even){ background: #CCCCCC; }
table tr:nth-child(odd) td:nth-child(odd){ background: #F2F2F2; }
table tr:nth-child(even) td:nth-child(odd){ background: #DDDDDD; }
table tr:nth-child(odd) td:nth-child(even){ background: #E5E5E5; }
div.column { width: 320px; float: left; }
div.first{ padding-right: 20px; border-right: 1px  grey solid; }
div.second{ margin-left: 30px; }
table{ margin-left: 20px; }
-->
</style>
</head>
<body>
'@

$HTMLEnd = @'
</div>
</body>
</html>
'@

# Date Format
$DateFormat      = Get-Date -Format "MM/dd/yyyy_HHmmss" 
#endregion Configuration

#region Help
<#
.SYNOPSIS
   Get-GeneralSystemReport.ps1
.DESCRIPTION

.PARAMETER
   <none>
.INPUTS
   <none>
.OUTPUTS
   <none>
.EXAMPLE
   Run stand alone
      Get-GeneralSystemReport.ps1
.LINK
   http://blogs.technet.com/b/exchange/archive/2011/01/18/3411844.aspx
#>
#endregion help

#region Functions
Function Get-DriveSpace() 
{
   Param (
   [string[]]$computers=@($env:computername)
   )

   $Title="Drive Report"

   #define an array for html fragments
   $fragments=@()

   #get the drive data
   $data=get-wmiobject -Class Win32_logicaldisk -filter "drivetype=3" -computer $computers

   #group data by computername
   $groups=$Data | Group-Object -Property SystemName

   #this is the graph character
   [string]$g=[char]9608 

   #create html fragments for each computer
   #iterate through each group object
           
   ForEach ($computer in $groups) {
       #define a collection of drives from the group object
       $Drives=$computer.group
       
       #create an html fragment
       $html=$drives | Select @{Name="Drive";Expression={$_.DeviceID}},
	   @{Name="Volume Name";Expression={$_.VolumeName}},
       @{Name='Size GB';Expression={$_.Size/1GB  -as [int]}},
       @{Name='Used GB';Expression={"{0:N2}" -f (($_.Size - $_.Freespace)/1GB) }},
       @{Name='Free GB';Expression={"{0:N2}" -f ($_.FreeSpace/1GB) }},
	   @{Name='% Free';Expression={"{0:N2}" -f (($_.FreeSpace/$_.Size)*100)}},
       @{Name="Usage";Expression={
         $UsedPer= (($_.Size - $_.Freespace)/$_.Size)*100
         $UsedGraph=$g * ($UsedPer/2)
         $FreeGraph=$g* ((100-$UsedPer)/2)
         #I'm using place holders for the < and > characters
         "xopenFont color=Redxclose{0}xopen/FontxclosexopenFont Color=Greenxclose{1}xopen/fontxclose" -f $usedGraph,$FreeGraph
       }} | ConvertTo-Html -Fragment 
       
       #replace the tag place holders. It is a hack but it works.
       $html=$html -replace "xopen","<"
       $html=$html -replace "xclose",">"
       
       #add to fragments
       $Fragments+=$html
       
       #insert a return between each computer
       $fragments+="<br>"
       
   } #foreach computer

   #write the result to a file
   Return $fragments
}

Function Get-HostUptime 
{
   param ([string]$ComputerName)
   $Uptime = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName
   $LastBootUpTime = $Uptime.ConvertToDateTime($Uptime.LastBootUpTime)
   $Time = (Get-Date) - $LastBootUpTime
   Return '{0:00} Days, {1:00} Hours, {2:00} Minutes, {3:00} Seconds' -f $Time.Days, $Time.Hours, $Time.Minutes, $Time.Seconds
}
#endregion functions

#region General System Report
$DriveSpaceReport = Get-DriveSpace $ServerName
# General System Info
$OS = (Get-WmiObject Win32_OperatingSystem -computername $ServerName).caption
$SystemInfo = Get-WmiObject -Class Win32_OperatingSystem -computername $ServerName | Select-Object Name, TotalVisibleMemorySize, FreePhysicalMemory
$TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB
$FreeRAM = $SystemInfo.FreePhysicalMemory/1MB
$UsedRAM = $TotalRAM - $FreeRAM
$RAMPercentFree = ($FreeRAM / $TotalRAM) * 100
$TotalRAM = [Math]::Round($TotalRAM, 2)
$FreeRAM = [Math]::Round($FreeRAM, 2)
$UsedRAM = [Math]::Round($UsedRAM, 2)
$RAMPercentFree = [Math]::Round($RAMPercentFree, 2)

$TopProcesses = Get-Process -ComputerName $ServerName | Sort WS -Descending | `
  Select ProcessName, Id, WS -First $ProccessNumToFetch | ConvertTo-Html -Fragment

# Services Report
$ServicesReport = @()
$Services = Get-WmiObject -Class Win32_Service -ComputerName $ServerName | `
  Where {($_.StartMode -eq "Auto") -and ($_.State -eq "Stopped")}

foreach ($Service in $Services) {
  $row = New-Object -Type PSObject -Property @{
        Name = $Service.Name
     Status = $Service.State
     StartMode = $Service.StartMode
  } 
  $ServicesReport += $row
}

$ServicesReport = $ServicesReport | ConvertTo-Html -Fragment
   
# Event Logs Report
$SystemEventsReport = @()
$SystemEvents = Get-EventLog -ComputerName $ServerName -LogName System -EntryType Error,Warning -Newest $EventNum
foreach ($event in $SystemEvents) {
   $row = New-Object -Type PSObject -Property @{
      TimeGenerated = $event.TimeGenerated
      EntryType = $event.EntryType
      Source = $event.Source
      Message = $event.Message
   }
   $SystemEventsReport += $row
}
      
$SystemEventsReport = $SystemEventsReport | ConvertTo-Html -Fragment

$ApplicationEventsReport = @()
$ApplicationEvents = Get-EventLog -ComputerName $ServerName -LogName Application -EntryType Error,Warning -Newest $EventNum
foreach ($event in $ApplicationEvents) {
   $row = New-Object -Type PSObject -Property @{
      TimeGenerated = $event.TimeGenerated
      EntryType = $event.EntryType
      Source = $event.Source
      Message = $event.Message
   }
   $ApplicationEventsReport += $row
}

$ApplicationEventsReport = $ApplicationEventsReport | ConvertTo-Html -Fragment

# Uptime
# Fetch the Uptime of the current system using our Get-HostUptime Function.
$SystemUptime = Get-HostUptime -ComputerName $ServerName

# Create HTML Report for the current System being looped through
$CurrentSystemHTML = ''
$CurrentSystemHTML += "<hr noshade size=3 width='100%'>"
$CurrentSystemHTML += "<div id='report'>"
$CurrentSystemHTML += "<p><h2>$Servername</p></h2>"
$CurrentSystemHTML += "<h3>System Info</h3>"
$CurrentSystemHTML += '<table class="list">'
$CurrentSystemHTML += '<tr>'
$CurrentSystemHTML += '<td>System Uptime</td>'
$CurrentSystemHTML += "<td>$SystemUptime</td>"
$CurrentSystemHTML += "</tr>"
$CurrentSystemHTML += "<tr>"
$CurrentSystemHTML += "<td>OS</td>"
$CurrentSystemHTML += "<td>$OS</td>"
$CurrentSystemHTML += "</tr>"
$CurrentSystemHTML += "<tr>"
$CurrentSystemHTML += "<td>Total RAM (GB)</td>"
$CurrentSystemHTML += "<td>$TotalRAM</td>"
$CurrentSystemHTML += "</tr>"
$CurrentSystemHTML += "<tr>"
$CurrentSystemHTML += "<td>Free RAM (GB)</td>"
$CurrentSystemHTML += "<td>$FreeRAM</td>"
$CurrentSystemHTML += "</tr>"
$CurrentSystemHTML += "<tr>"
$CurrentSystemHTML += "<td>Percent free RAM</td>"
$CurrentSystemHTML += "<td>$RAMPercentFree</td>"
$CurrentSystemHTML += "</tr>"
$CurrentSystemHTML += "</table>"
$CurrentSystemHTML += "<h3>Disk Info</h3>"
$CurrentSystemHTML += "$DriveSpaceReport"
$CurrentSystemHTML += "<br></br>"
$CurrentSystemHTML += "<table class='normal'>"
$CurrentSystemHTML += "$DiskInfo</table>"
$CurrentSystemHTML += "<br></br>"
$CurrentSystemHTML += "<div class='first column'>"
$CurrentSystemHTML += "<h3>System Processes - Top $ProccessNumToFetch Highest Memory Usage</h3>"
$CurrentSystemHTML += "<p>The following $ProccessNumToFetch processes are those consuming the highest amount of Working Set (WS) Memory (bytes) on $Servername</p>"
$CurrentSystemHTML += "<table class='normal'>"
$CurrentSystemHTML += "$TopProcesses</table>"
$CurrentSystemHTML += "</div>"
$CurrentSystemHTML += "<div class='second column'>"
$CurrentSystemHTML += "<h3>System Services - Automatic Startup but not Running</h3>"
$CurrentSystemHTML += "<p>The following services are those which are set to Automatic startup type, yet are currently not running on $Servername</p>"
$CurrentSystemHTML += "<table class='normal'>"
$CurrentSystemHTML += "$ServicesReport"
$CurrentSystemHTML += "</table>"
$CurrentSystemHTML += "</div>"
$CurrentSystemHTML += "<h3>Events Report - The last $EventNum System/Application Log Events that were Warnings or Errors</h3>"
$CurrentSystemHTML += "<p>The following is a list of the last $EventNum <b>System log</b> events that had an Event Type of either Warning or Error on $Servername</p>"
$CurrentSystemHTML += "<table class='normal'>"
$CurrentSystemHTML += "$SystemEventsReport</table>"
$CurrentSystemHTML += "<p>The following is a list of the last $EventNum <b>Application log</b> events that had an Event Type of either Warning or Error on $Servername</p>"
$CurrentSystemHTML += "<table class='normal'>"
$CurrentSystemHTML += "$ApplicationEventsReport</table>"

# Add the current System HTML Report into the final HTML Report body
$HTMLMiddle += $CurrentSystemHTML

# Assemble the final report from all our HTML sections
$HTMLmessage = $HTMLHeader + $HTMLMiddle + $HTMLEnd

if ($SendMail)
{
	$HTMLmessage = $HTMLmessage | Out-String
	$email= 
	@{
		From = $EmailSender
		To = $EmailRecipient
#		CC = "EMAIL@EMAIL.COM"
		Subject = "General Server Report - $ServerName"
		SMTPServer = $EmailRelay
		Body = $HTMLmessage
		Encoding = ([System.Text.Encoding]::UTF8)
		BodyAsHTML = $true
	}
	Send-MailMessage @email
	Sleep -Milliseconds 200
}
elseif ($SaveReport)
{
	$HTMLMessage | Out-File $ReportName
}
else
{
   Return $HTMLmessage
}
#endregion General System Report


