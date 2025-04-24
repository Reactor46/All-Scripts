$LogTime = Get-Date -Format yyyy-MM-dd_hh-mm
$LogFile = "c:\healtcheck\Logs\DailyHealthCheck-$LogTime.rtf"
[int]$EventNum = 3
if(!(Test-Path -Path 'c:\healtcheck\Logs')) {mkdir 'c:\healtcheck\Logs'}
if(!(Test-Path -Path 'c:\healtcheck\Reports')) {mkdir 'c:\healtcheck\Reports'}

# Add SharePoint PowerShell Snapin

if ( (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null ) {
    Add-PSSnapin Microsoft.SharePoint.Powershell
}
import-module WebAdministration

$scriptBase = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
Set-Location $scriptBase

write-host "TESTING FOR LOG FOLDER EXISTENCE" -fore blue
$TestLogFolder = test-path -path $scriptbase\Logs

	

#moving any .rtf files in the scriptbase location
$FindRTFFile = Get-ChildItem $scriptBase\*.* -include *.rtf
if($FindRTFFile)
{
	write-host "Some old log files are found in the script location" -fore blue
	write-host "Moving old log files into the Logs folder" -fore blue
	foreach($file in $FindRTFFile)
		{
			move-item -path $file -destination $scriptbase\logs
		}
	write-host "Old log files moved successfully" -fore blue
}

#start-transcript $logfile

$global:timerServiceName = "SharePoint Timer Service"
$global:timerServiceInstanceName = "Microsoft SharePoint Foundation Timer"

# Get the local farm instance
[Microsoft.SharePoint.Administration.SPFarm]$farm = [Microsoft.SharePoint.Administration.SPFarm]::get_Local()

Function SharePointServices([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	Write-Host ""
	write-host "Generating SharePoint services report" -fore Magenta
	
	$output = $scriptbase + "\Reports\" + "SharePointServices.csv"
	"ServiceName" + "," + "ServiceStatus" + "," + "StartMode" + "," + "MachineName" | Out-File -Encoding Default -FilePath $Output;

	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			    [string]$serverName = $server.Name
				write-host "Generating SP services report for server" $serverName -fore blue
				$Monitors = "SPAdminV4" , "SPTimerV4" , "SPTraceV4" , "SPSearch4" , "OSearch16"
				foreach($monitor in $Monitors){
				
				    $service = Get-Service -ComputerName $serverName -Name $monitor -ea silentlycontinue |Sort-Object -Descending status
                
                    $startup = Get-WmiObject -Class Win32_Service -Property StartMode -ComputerName $serverName -Filter "Name='$monitor'" 

					$service.displayname + "," + $service.status + "," + $startup.StartMode + "," + $service.MachineName | Out-File -Encoding Default  -Append -FilePath $Output;
				}
				write-host "SP services report generated" -fore green
			}
		}
	}

} 

Function IISWebsite([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	Write-Host ""
	write-host "Generating IIS website report" -fore Magenta
	
	$output = $scriptbase + "\Reports\" + "IISWebsite.csv"
	"WebSiteName" + "," + "WebsiteID" + "," + "WebSiteState" + "," + "Server" | Out-File -Encoding Default -FilePath $Output;

	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			              [string]$serverName = $server.Name
				write-host "Generating IIS website report for server" $serverName -fore blue

				$status = ""
				#$Sites = gwmi -namespace "root/MicrosoftIISv2" -class IIsApplicationPoolSetting -ComputerName $serverName -Authentication PacketPrivacy -Impersonation Impersonate |  Where-Object { ($_.name -notcontains "Default Web Site") }
				$Sites = gwmi -namespace "root\webadministration" -Class site -ComputerName $serverName -Authentication PacketPrivacy -Impersonation Impersonate |  Where-Object { ($_.name -notcontains "Default Web Site") }
                foreach($site in $sites)
				{
					if($site.getstate().returnvalue -eq 1)
					{
						$status = "Started"
					}
					else
					{
						$status = "Stopped"
					}
				

					$site.name + "," + $site.ID + "," + $Status + "," + $serverName | Out-File -Encoding Default  -Append -FilePath $Output;
				}
				write-host "IIS website report generated" -fore green
			}
		}
	}

}


Function AppPoolStatus([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	Write-Host ""
	write-host "Generating AppPool status report" -fore Magenta
	
	$output = $scriptbase + "\Reports\" + "AppPoolStatus.csv"
	"AppPoolName" + "," + "Status" + "," + "Server" | Out-File -Encoding Default -FilePath $Output;

	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			              [string]$serverName = $server.Name
				write-host "Generating AppPool status report for server" $serverName -fore blue

				$status = ""
				#$AppPools = gwmi -namespace "root/MicrosoftIISv2" -class IIsApplicationPoolSetting -ComputerName $serverName -Authentication PacketPrivacy -Impersonation Impersonate |Where-Object { ($_.name -notcontains "SharePoint Web Services Root")}
				$AppPools = gwmi -namespace "root\webadministration" -Class applicationpool -ComputerName $serverName -Authentication PacketPrivacy -Impersonation Impersonate |Where-Object { ($_.name -notcontains "SharePoint Web Services Root")}
                foreach($AppPool in $AppPools )
				{
					if($AppPool.getstate().returnvalue -eq 1)
					{
						$status = "Started"
					}
					else
					{
						$status = "Stopped"

					}
				

					$AppPool.name + "," + $Status + "," + $serverName| Out-File -Encoding Default  -Append -FilePath $Output;
				}
				write-host "AppPool status report generated" -fore green
			}
		}
	}

}


Function DiskSpace([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	Write-Host ""
	write-host "Generating Disk space report" -fore Magenta
	
	$output = $scriptbase + "\Reports\" + "DiskSpace.csv"
	"Computer Name" + "," + "Drive" + "," + "Size in (GB)" + "," + "Free Space in (GB)" + "," + "Percentage"+ "," + "Status"  | Out-File -Encoding Default -FilePath $Output;

	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			              [string]$serverName = $server.Name
				write-host "Generating disk space report for server" $serverName -fore blue

				$drives = Get-WmiObject -ComputerName $serverName Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3}  

 				foreach($drive in $drives)  
	 			{  
					$id = $drive.DeviceID  

 					$size = [math]::round($drive.Size / 1073741824, 2)  

 					$free = [math]::round($drive.FreeSpace  / 1073741824, 2)  

 					$pct = [math]::round($free / $size, 2) * 100  

					if ($pct -lt 10) 
					{ 
						$pct = $pct.ToString() + "% *" 
                        $stat = "Not OK"

					}  
					else
					{
						$pct = $pct.ToString() + " %" 
                        $stat = "OK"
					}  

   					$serverName + "," + $id + "," + $size + "," + $free + "," + $pct + "," + $stat  | Out-File -Encoding Default  -Append -FilePath $Output;
					$pct = 0   
				}
				write-host "Disk space report generated" -fore green
			}
		}
	}

}


Function HealthAnalyserReports()
{
	
	write-host ""
	write-host "Generating health analyser report" -fore magenta

	$output = $scriptbase + "\Reports\" + "HealthAnalyser.csv"	
	"Severity" + "," + "Category" + "," + "Modified" + "," + "Failing servers" + "," + "Failing services"  | Out-File -Encoding Default -FilePath $Output;

	$ReportsList = [Microsoft.SharePoint.Administration.Health.SPHealthReportsList]::Local
	$Items = $ReportsList.items | where {	

		if($_['Severity'] -eq '1 - Error')
		{
			#write-host $_['Name']
			#write-host $_['Severity']
            $server_lok1 = $_['Failing Servers']  
            $server_lok1 = $server_lok1 -replace "`n","";
			$_['Severity'] + "," + $_['Category'] + "," + $_['Modified'] + "," + $server_lok1 + "," + $_['Failing Services']  | Out-File -Encoding Default  -Append -FilePath $Output;            
#			$_['Severity'] + "," + $_['Category'] + "," + $_['Modified'] + "," + $_['Failing Servers'] + "," + $_['Failing Services']  | Out-File -Encoding Default  -Append -FilePath $Output;
		}
	}
	write-host "Health analyser report generated" -fore green
}



Function CPUUtilization([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	Write-Host ""
	write-host "Generating CPU utilization report" -fore Magenta
	
	$output = $scriptbase + "\Reports\" + "CPUUtilization.csv"
	"ServerName" + "," + "DeviceID" + "," + "LoadPercentage" + "," + "Status" | Out-File -Encoding Default -FilePath $Output;

	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			              [string]$serverName = $server.Name
				write-host "Generating CPU utilization report for server" $serverName -fore blue
				$CPUDataCol = Get-WmiObject -Class Win32_Processor -ComputerName $ServerName 
				foreach($Data in $CPUDataCol)
				{
					$serverName + "," + $Data.DeviceID + "," + $Data.loadpercentage + "," + $Data.status | Out-File -Encoding Default  -Append -FilePath $Output;

                }
				write-host "CPU utilization report generated" -fore green
			}
		}
	}

}


Function MemoryUtilization([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	Write-Host ""
	write-host "Memory utilization report" -fore Magenta
	
	$output = $scriptbase + "\Reports\" + "MemoryUtilization.csv"
	"ServerName" + "," + "FreePhysicalMemory" + "," + "TotalVisibleMemorySize" + "," + "Status" | Out-File -Encoding Default -FilePath $Output;

	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			              [string]$serverName = $server.Name
				write-host "Generating memory utilization report for server" $serverName -fore blue
				$MemoryCol = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ServerName
				foreach($Data in $MemoryCol)
				{
					$serverName + "," + $Data.FreePhysicalMemory + "," + $Data.TotalVisibleMemorySize + "," + $Data.status | Out-File -Encoding Default  -Append -FilePath $Output;

                }
				write-host "Memory utilization report generated" -fore green
			}
		}
	}

}

Function Hotfixes([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	$output = $scriptbase + "\Reports\" + "hotfix.csv"
	"Hotfix ID" + "," + "Description" + "," + "Installation date" + "," + "Server" | Out-File -Encoding Default -FilePath $Output;


	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			              [string]$serverName = $server.Name
				
			$lastdate= (get-date).AddDays(-180)

                $hotfixes = Get-HotFix  -ComputerName $serverName -ea silentlycontinue | where { $_.installedon -ge $lastdate} |Sort-Object -Descending status

				foreach ($hotfix in $hotfixes) {

				$hotfix.HotFixID + "," + $hotfix.Description + "," + $hotfix.InstalledOn  + "," + $hotfix.PSComputerName | Out-File -Encoding Default  -Append -FilePath $Output;

                }				
				write-host "Hotfix report generated" -fore green
			}
		}
	}

}

Function AppEvents([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	Write-Host ""
	write-host "SharePoint server application events" -fore Magenta
	
    $output = $scriptbase + "\Reports\" + "AppEvents.csv"
	"Server" + "," + "TimeGenerated" + "," + "EntryType" + "," + "Message" | Out-File -Encoding Default -FilePath $Output;

	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			    [string]$serverName = $server.Name
                write-host "Collecting application events report for server" $serverName -fore blue
				$AppEventCol = Get-EventLog -ComputerName $serverName -LogName Application -EntryType Error,Warning -Newest $EventNum
				foreach($aEvent in $AppEventCol)
				{
					$serverName + "," + $aEvent.TimeGenerated + "," + $aEvent.EntryType + "," + $aEvent.Message.Substring(0,$aEvent.Message.Length) | Out-File -Encoding Default -Append -FilePath $Output;

                }
				write-host "Application events report generated" -fore green
			}
		}
	}

}

Function SysEvents([Microsoft.SharePoint.Administration.SPFarm]$farm)
{

	Write-Host ""
	write-host "SharePoint server system events" -fore Magenta
	
    $output = $scriptbase + "\Reports\" + "SysEvents.csv"
	"Server" + "," + "TimeGenerated" + "," + "EntryType" + "," + "Source" + "," + "Message" | Out-File -Encoding Default -FilePath $Output;

	foreach($server in $farm.Servers)
    	{
		foreach($instance in $server.ServiceInstances)
        		{
			# If the server has the timer service then stop the service
		              if($instance.TypeName -eq $timerServiceInstanceName)
			{
			    [string]$serverName = $server.Name
                write-host "Collecting system events report for server" $serverName -fore blue
				$SysEventCol = Get-EventLog -ComputerName $serverName -LogName System -EntryType Error,Warning -Newest $EventNum
				foreach($sEvent in $SysEventCol)
				{
					$serverName + "," + $sEvent.TimeGenerated + "," + $sEvent.EntryType + "," + $sEvent.Source + "," + $sEvent.Message.Substring(0,$sEvent.Message.Length) | Out-File -Encoding Default  -Append -FilePath $Output;
				
                }
				write-host "System events report generated" -fore green
			}
		}
	}

}

<#

Function Httprequest 
{
	Write-Host ""
	write-host "Http Respond Report" -fore Magenta    
    
    $WebApp = Get-SPWebApplication -IncludeCentralAdministration
    $URLs =  $WebApp |Select url
  
	$output = $scriptbase + "\Reports\" + "HttpResponse.csv"
    "URL" + "," + "HTTP status" + "," + "HTTP Description" + "," + "Error ?" | Out-File -Encoding Default -FilePath $output;
    write-host "Generating Http respond report" $serverName -fore blue
    foreach ($URL in $URLs)
        {
        $error=$null
            try {
                $request = Invoke-WebRequest -Uri $URL -MaximumRedirection 5 -ErrorAction Ignore
                } 
            catch [System.Net.WebException] 
                {
                $request = $_.Exception.Message
                $error =  $_  
                }
         $URL + "," + $request.StatusCode  + "," + $request.StatusDescription  + "," +  $error.Exception.Message | Out-File -Encoding Default -Append  -FilePath $output;
         }
    write-host "Http Respond Report generated" -fore green
}

#>

#######################Function to combine multiple CSV files into single excel sheet with seperated tabs for each CSV#########################

Function Release-Ref ($ref) 
{
	([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
	[System.__ComObject]$ref) -gt 0)
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers() 
}


#################################################################################################################################################


##########Calling Functions#################
SharePointServices $farm
IISWebsite $farm
AppPoolStatus $farm
DiskSpace $farm
HealthAnalyserReports
CPUUtilization $farm
MemoryUtilization $farm
Hotfixes $farm
AppEvents $farm
SysEvents $farm
#Httprequest


write-host ""
write-host "Combining all CSV files into single file" -fore blue

$Head = 
@"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "https://protect-de.mimecast.com/s/EGMPCoZ3nxC5BAnAS1EA-U?domain=w3.org">
<html><head>
<style type="text/css">
<!--
body {
	font-family: Verdana;
	margin: 0;
	padding: 0;
}

table {
    border: 1px collapse white;
    color: #000000;
    font-size: 80%;
}

th, td {
    border: 1px solid black;
    text-align: left;
    padding: 1px;
}

tr:nth-child(even){background-color: #f2f2f2}

th {
    background-color: #222930;
    color: white;
}

h1{ /*HEADER*/
    clear: both; 
    font-size: 200%;
    background: #222930; 
    color: #ffffff;
    padding: 20px;
    text-transform: uppercase;
}

h2{  /*SEKCJE*/
    clear: both; 
    font-size: 130%;
    background: #FFD500; /*TEN KOLOR - MOZNA DOPASOWAC DO KONTRAKTU*/
    color: #000000;
    padding: 5px;
    text-transform: uppercase;
}

br.clear {
	clear: both;
}

-->
</style>
</head>
<body>
"@

#Cell Color - Logic
$StatusColor = @{
"Not OK" = ' bgcolor="Red">"Not OK"<'; 
"Disabled" = ' bgcolor="blue">Disabled<'; 
"Manual" = ' bgcolor="blue">Manual<'; 
"Stopped" = ' bgcolor="Red">Stopped<';
"Not Started" = ' bgcolor="Red">Stopped<';
}


$serwisy = Get-Content c:\healtcheck\Reports\SharePointServices.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>SharePoint Services</h2>" |Out-String  
$iisy = Get-Content c:\healtcheck\Reports\IISWebsite.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>IIS Website Status</h2>" |Out-String   
$poole= Get-Content c:\healtcheck\Reports\AppPoolStatus.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>Application Pool Status</h2>" |Out-String 
$zdrowko= Get-Content c:\healtcheck\Reports\HealthAnalyser.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>Health Analyzer</h2>" |Out-String 
$space = Get-Content c:\healtcheck\Reports\DiskSpace.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>Disks Space</h2>" |Out-String 
$proc= Get-Content c:\healtcheck\Reports\CPUUtilization.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>CPU Utilization</h2>" |Out-String 
$pam= Get-Content c:\healtcheck\Reports\MemoryUtilization.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>Memory Utilization</h2>" |Out-String 
$fix= Get-Content c:\healtcheck\Reports\Hotfix.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>Htofix Status</h2>" |Out-String 
$appev = Get-Content c:\healtcheck\Reports\AppEvents.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>Application Events</h2>" |Out-String 
$sysev = Get-Content c:\healtcheck\Reports\SysEvents.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>System Events</h2>" |Out-String 
#$httpstat = Get-Content c:\healtcheck\Reports\HttpResponse.csv |ConvertFrom-Csv | ConvertTo-HTML -Head $Head -Body "<h2>Http Response</h2>" |Out-String 

$StatusColor.Keys | foreach { $serwisy = $serwisy -replace ">$_<",($StatusColor.$_) }
$StatusColor.Keys | foreach { $iisy = $iisy -replace ">$_<",($StatusColor.$_) }
$StatusColor.Keys | foreach { $poole = $poole -replace ">$_<",($StatusColor.$_) }
$StatusColor.Keys | foreach { $space = $space -replace ">$_<",($StatusColor.$_) }
$StatusColor.Keys | foreach { $proc = $proc -replace ">$_<",($StatusColor.$_) }
$StatusColor.Keys | foreach { $pam = $pam -replace ">$_<",($StatusColor.$_) }

#ConvertTo-HTML -head $Head -PostContent $serwisy, $iisy, $poole, $httpstat, $zdrowko, $space, $proc, $pam, $fix, $appev, $sysev -PreContent '<h1>SP PROD: Full Health Check</h1>'|Out-File D:\script\HealthCheck\Reports\TableHTML.html 
ConvertTo-HTML -head $Head -PostContent $serwisy, $iisy, $poole, $zdrowko, $space, $proc, $pam, $fix, $appev, $sysev -PreContent '<h1>SP PROD: Full Health Check</h1>'|Out-File c:\healtcheck\Reports\TableHTML.html 



#Get-Item $scriptbase\*.csv | ConvertCSV-ToExcel -output "DailyMonitoringReports.xlsx"


write-host ""
write-host "SCRIPT COMPLETED" -fore green

#stop-transcript

#Send report to the recipient email address
<#
$file= "c:\healtcheck\Reports\TableHTML.html"
$att = new-object Net.Mail.Attachment($file)
$smtpServer = ""
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg = new-object Net.Mail.MailMessage
$msg.From = ""
$recipient = "piotr.jablonski@atos.net"
$msg.To.Add($recipient)
$msg.Subject = "PROD Full Health Check"
$msg.Body = "Please find related information in log file attached"
#$msg.Body = Get-Content $file
$msg.Attachments.Add($att)
$msg.IsBodyHTML = $true
$smtp.Send($msg)
$att.Dispose()
#>
