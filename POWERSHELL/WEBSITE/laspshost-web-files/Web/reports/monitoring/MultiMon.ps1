#######################################################################################
### Multiple Windows Services Monitor (Powershell)
###
### Params: ServiceNames
###   ServiceNames = List of Services separated by pipe
###
### Example: SCRIPT.PS1 wuauserv|SNMP-Trap|gupdate|bthserv
###
### Return Codes:
###   OK = 0
###   DOWN = 1
###   WARNING = 2
###   ERROR = 3
###
#######################################################################################

# Variables
$ServiceNames = $args.get(0).split('|')
$RunResults = New-Object System.Collections.Generic.List[System.Object]
$resultcode = 0
$total = 0

# Functions
Function CheckServiceExist ($ServiceName)
{
	if (!(Get-Service -Name $ServiceName -ErrorAction SilentlyContinue))
	{
		$RunResults.Add("WARNING: $ServiceName does not exist")
		return $false
	}else{return $true}
}

Function CheckServiceState ($ServiceName)
{
	$ServiceState = (Get-Service -Name $ServiceName).Status
	if(!($ServiceState -eq 'Running')){
		$RunResults.Add("CRITICAL: $ServiceName -&gt; $ServiceState")
		}else{$RunResults.Add("OK: $ServiceName -&gt; $ServiceState")}
}

# Main
foreach($ServiceName in $ServiceNames)
{
	$srv_exist = (CheckServiceExist -ServiceName $ServiceName)
	if ($srv_exist){
		CheckServiceState -ServiceName $ServiceName	
	}
}

if($RunResults -match "WARNING:"){$resultcode = 2}
if($RunResults -match "CRITICAL:"){$resultcode = 3}
$count_ok = ($RunResults | select-string -pattern "OK:").length
$count_warning = ($RunResults | select-string -pattern "WARNING:").length
$count_critical = ($RunResults | select-string -pattern "CRITICAL:").length
$total_errors = $count_warning + $count_critical
$total_results = $RunResults.Count

$lines = @()
for ( $i = 1 ; $i -le $RunResults.Count; $i++ ) 
{
  $lines += "&lt;br/&gt;"
  $lines += $RunResults[$RunResults.Count - $i]
}

# Output
write-host "Statistic: $total_errors"

if ($resultcode -eq 2)
 {
 write-host "Message: $count_ok of $total_results monitored services running $lines"
 exit $resultcode
 }
 
if ($resultcode -eq 3)
 {
 write-host "Message: $count_ok of $total_results monitored services running $lines"
 exit $resultcode
 }
 
if ($resultcode -eq 0)
 {
 write-host "Message: $count_ok of $total_results monitored services running $lines"
 exit $resultcode
 }