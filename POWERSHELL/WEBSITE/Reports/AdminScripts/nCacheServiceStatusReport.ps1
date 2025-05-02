# Variables below are for editing the style of the HTML thats generated at the end.
$a = '<header><meta http-equiv="refresh" content="30" ></header>'
$a = $a + "<style>"
$a = $a + "BODY{background-color:gainsboro;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:limegreen}"
$a = $a + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:whitesmoke}"
$a = $a + "</style>"
# Text to be displayed BEFORE the table. Note: the 'every five minutes' reference below is due to how often the task for this script will run.
$Pre = "Status refreshes every 30 seconds."
# Text to be displayed AFTER the table
$Post = "Missing Application?"
# Replace 'servername' with the server that is hosting the services.
# Replace 'serviceName' with the name of the service(s) you want to include on this report.



function Get-nCacheMemUsage {
Get-Process -ComputerName (GC C:\inetpub\AdminScripts\Configs\Servers.txt) -Name Alachisoft.NCache.Service |
    Select @{Expression = {$_.MachineName};Label="Server"}, @{Expression={[String]([int]($_.WS/1MB))+" MB"};Label="Memory Usage"} |
        Sort-Object  Server | ConvertTo-Html -title "Critical Application Status" -body "<H2>Critical Application Status</H2>" -head $a -PreContent $Pre -PostContent $Post |
            Set-Content C:\inetpub\wwwroot\Report\ncache_status.htm


            Start-Sleep -Seconds 30
}

 
while($true){
 Get-nCacheMemUsage
 }