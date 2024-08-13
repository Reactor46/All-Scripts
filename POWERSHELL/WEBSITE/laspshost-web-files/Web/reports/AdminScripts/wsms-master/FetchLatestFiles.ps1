#########################################################
#                                                       #
# Fetching latest Monitoring Windows Server Scripting   #
#                                                       #
#########################################################

$url = "https://raw.githubusercontent.com/tarlety/wsms/master/MonitoringWindowsServerScripting.ps1"
$path = $PSScriptRoot + "\MonitoringWindowsServerScripting.ps1"

$client = New-Object System.Net.WebClient
$client.DownloadFile($url, $path)
