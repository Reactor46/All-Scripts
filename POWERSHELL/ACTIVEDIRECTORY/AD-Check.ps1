############################## AD Health Check #############################

####### This will provide HTML report for AD health status  ################


############################################################################


$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 2px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 2px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
Title of my Report
</title>
"@


###### Function to Check Service Status ######

Function Getservicestatus($service, $server)
{
	$st = Get-service -computername $server | where-object { $_.name -eq $service }
	if($st)
	{$servicestatus= $st.status}
	else
	{$servicestatus = "Not found"}
	
	Return $servicestatus
}



####### Find Domain Controllers in Forest  ########

$Forest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()

[string[]]$computername = $Forest.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 


####### Check Server Availability ###########

$report1= @()
foreach ($server in $computername){
$temp1 = "" | select server, pingstatus
if ( Test-Connection -ComputerName $server -Count 1 -ErrorAction SilentlyContinue ) {
$temp1.pingstatus = "Pinging"
}
else {
$temp1.pingstatus = "Not pinging"
}
$temp1.server = $server
$report1+=$temp1
}

$b = $report1 | select server, pingstatus  | ConvertTo-HTML  -Head $Header -As Table -PreContent "<h2>Server Availability</h2>" | Out-String



########## Check Service Status ####################

$report = @()

foreach ($server in $computername){
$temp = "" | select server, NTDS, DNS, DFSR, netlogon, w32Time
$temp.server = $server

$temp.NTDS = Getservicestatus -service "NTDS" -server $server
$temp.DNS = Getservicestatus -service "DNS" -server $server
$temp.DFSR = Getservicestatus -service "DFSR" -server $server
$temp.netlogon = Getservicestatus -service "netlogon" -server $server
$temp.w32Time = Getservicestatus -service "w32Time" -server $server
$report+=$temp
}

$b+= $REPORT | select server, NTDS, DNS, DFSR, netlogon, w32Time | ConvertTo-HTML -Head $Header -As Table -PreContent "<h2>Service Status</h2>" | Out-String


add-type -AssemblyName microsoft.visualbasic 
$strings = "microsoft.visualbasic.strings" -as [type] 


######### Check netLogon Status #############

$report = @()
foreach ($server in $computername){
$temp = "" | select server, SysvolTest
$temp.server = $server
$svt = dcdiag /test:netlogons /s:$server
if($strings::instr($svt, "passed test NetLogons")){$temp.SysvolTest = "Passed"}
else
{$temp.SysvolTest = "Failed"}
$report+=$temp
}
$b+= $REPORT | select server, SysvolTest | ConvertTo-HTML -Fragment -As Table -PreContent "<h2>NetLogon Test</h2>" | Out-String


######## Test Replication Status #############


$workfile = \\DC1\c$\Windows\System32\repadmin.exe /showrepl * /csv 
$results = ConvertFrom-Csv -InputObject $workfile 
 
 
$results = $results 
 

    $results = $results | select "Source DSA", "Naming Context", "Destination DSA" ,"Number of Failures", "Last Failure Time", "Last Success Time", "Last Failure Status"
    $b+= $results | select "Source DSA", "Naming Context", "Destination DSA" ,"Number of Failures", "Last Failure Time", "Last Success Time", "Last Failure Status" | ConvertTo-HTML -Head $Header -As Table -PreContent "<h2>Replication Status</h2>" | Out-String    






$head = @'
<style>
body { font-family:Tahoma;
       font-size:12pt; }
td, th { border:1px solid black; 
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px }
table { margin-left:50px; }
</style>
'@
 
$s = ConvertTo-HTML -head $head -PostContent $b -Body "<h1>Active Directory Checklist</h1>" | Out-string


$emailFrom = "CampusADHealth@test.com" 
$emailTo = "sharmamk@test.com"

$smtpserver= "smtp.test.com" 
$smtp=new-object Net.Mail.SmtpClient($smtpServer)

$msg = new-object Net.Mail.MailMessage
$msg.From = $emailFrom
$msg.To.Add($emailTo)
$msg.IsBodyHTML = $true
$msg.subject="Active Directory Health Check Report" 
$msg.Body = $s
$smtp.Send($msg)


