#############################################################################
#       Authors: Mahesh Sharma/ Vikas Sukhija 
#       Reviewer: Vikas SUkhija      
#       Date: 06/10/2013
#		Modified:06/19/2013 - made it to run from any path
#       Modified:02/09/2014 - started modifying it for exchange 2010
#       Modified:02/18/2014 - modified to include all mailox servers in test mailflow
#       Modified:05/22/2014 - added activation prefrence
#		Modified:09/09/2014 - included DAG DB backups status
#		Modified:04/18/2015 - Modified to work in Exchange 2010/2013 coexistence Enviornment
#		Modified:07/12/2015 - Updated to show yellow indicators if queue length increases 50
#		Modified:12/12/2015 - Rmoev the Red status for content index where index service is disabled 
#		Modified:05/08/2017 - Adding eamil ALert on db, index & backup
#       Description: ExChange Health Status
#############################################################################

########################### Add Exchange Shell##############################

If ((Get-PSSnapin | where { $_.Name -match "Microsoft.Exchange.Management.PowerShell.E2010" }) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}

#############Email function##############
################Email Function#####################

function Send-Email
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		$From,
		[Parameter(Mandatory = $true)]
		[array]$To,
		[array]$bcc,
		[array]$cc,
		$body,
		$subject,
		$attachment,
		[Parameter(Mandatory = $true)]
		$smtpserver
	)
	
	$message = new-object System.Net.Mail.MailMessage
	$message.From = $from
	if ($To -ne $null)
	{
		$To | ForEach-Object{
			$to1 = $_
			$to1
			$message.To.Add($to1)
		}
	}
	if ($cc -ne $null)
	{
		$cc | ForEach-Object{
			$cc1 = $_
			$cc1
			$message.CC.Add($cc1)
		}
	}
	if ($bcc -ne $null)
	{
		$bcc | ForEach-Object{
			$bcc1 = $_
			$bcc1
			$message.bcc.Add($bcc1)
		}
	}
	$message.IsBodyHtml = $True
	if ($subject -ne $null)
	{
		$message.Subject = $Subject
	}
	if ($attachment -ne $null)
	{
		$attach = new-object Net.Mail.Attachment($attachment)
		$message.Attachments.Add($attach)
	}
	if ($body -ne $null)
	{
		$message.body = $body
	}
	$smtp = new-object Net.Mail.SmtpClient($smtpserver)
	$smtp.Send($message)
}
#############Email ALerts Switch/report & Variables###########
$ALert = "Yes"
$htmlreporting = "Yes"
$smtphost = "smtp.labtest.com"
$from = "Exchange2010Status@labtest.com"
$from1 = "SysMonitoring@labtest.com"
$to = "Vikass@labtest.com"
$hrs = (get-date).Addhours(-24)

###########################Define Variables################

$reportpath = ".\2010Report.htm"

if ((test-path $reportpath) -like $false)
{
	new-item $reportpath -type file
}
##################HTml Report Content##########
$report = $reportpath

Clear-Content $report
Add-Content $report "<html>"
Add-Content $report "<head>"
Add-Content $report "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $report '<title>Exchange Status Report</title>'
add-content $report '<STYLE TYPE="text/css">'
add-content $report  "<!--"
add-content $report  "td {"
add-content $report  "font-family: Tahoma;"
add-content $report  "font-size: 11px;"
add-content $report  "border-top: 1px solid #999999;"
add-content $report  "border-right: 1px solid #999999;"
add-content $report  "border-bottom: 1px solid #999999;"
add-content $report  "border-left: 1px solid #999999;"
add-content $report  "padding-top: 0px;"
add-content $report  "padding-right: 0px;"
add-content $report  "padding-bottom: 0px;"
add-content $report  "padding-left: 0px;"
add-content $report  "}"
add-content $report  "body {"
add-content $report  "margin-left: 5px;"
add-content $report  "margin-top: 5px;"
add-content $report  "margin-right: 0px;"
add-content $report  "margin-bottom: 10px;"
add-content $report  ""
add-content $report  "table {"
add-content $report  "border: thin solid #000000;"
add-content $report  "}"
add-content $report  "-->"
add-content $report  "</style>"
Add-Content $report "</head>"
Add-Content $report "<body>"
add-content $report  "<table width='100%'>"
add-content $report  "<tr bgcolor='Lavender'>"
add-content $report  "<td colspan='7' height='25' align='center'>"
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>DAG Active Manager</strong></font>"
add-content $report  "</td>"
add-content $report  "</tr>"
add-content $report  "</table>"

add-content $report  "<table width='100%'>"
Add-Content $report  "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='10%' align='center'><B>Identity</B></td>"
Add-Content $report  "<td width='5%' align='center'><B>PrimaryActiveManager</B></td>"
Add-Content $report  "<td width='20%' align='center'><B>OperationalMachines</B></td>"


Add-Content $report "</tr>"

##############################Get ALL DAG's##################################
$inputdag = @()

$indag = Get-DatabaseAvailabilityGroup

foreach ($dg in $indag)
{
	$mem = $dg.Servers
	foreach ($m in $mem)
	{
		if ((Get-ExchangeServer $m.Name).AdminDisplayVersion -like "*14.3*")
		{
			$inputdag += $dg.Name
		}
	}
}

$inputdag = $inputdag | select -uniq

################################################################################################################
################################################################################################################


$dagList = $inputdag
$TestMailFlow = Get-ExchangeServer | where{ $_.ServerRole -like "*Mailbox*" }

$report = $reportpath


##########################################################################################################
##############################################Check PAM###################################################

foreach ($dag in $dagList)
{
	
	$FullStatus = Get-DatabaseAvailabilityGroup -Status $dag
	
	Foreach ($status in $Fullstatus)
	{
		
		
		$Identity = $status.identity
		$PrimaryActiveManager = $status.PrimaryActiveManager
		$Servers = $status.Servers
		Add-Content $report "<tr>"
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>"
		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$PrimaryActiveManager</B></td>"
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$Servers</B></td>"
		Add-Content $report "</tr>"
		
		
	}
	
	
}

##################################################################################################################
############################################## Mailbox Database Status ###########################################


add-content $report  "<tr bgcolor='Lavender'>"
add-content $report  "<td colspan='7' height='25' align='center'>"
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Mailbox Database Status</strong></font>"
add-content $report  "</td>"
add-content $report  "</tr>"

add-content $report  "</tr>"
add-content $report  "</table>"
add-content $report  "<table width='100%'>"
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='25%' align='center'><B>databaseName</B></td>"
Add-Content $report "<td width='25%' align='center'><B>Status</B></td>"
Add-Content $report "<td width='25%' align='center'><B>ActiveCopy</B></td>"
Add-Content $report  "<td width='25%' align='center'><B>CopyQueuelength</B></td>"
Add-Content $report  "<td width='25%' align='center'><B>ReplayQueueLength</B></td>"
Add-Content $report  "<td width='25%' align='center'><B>LastInspectedLogTime</B></td>"
Add-Content $report  "<td width='25%' align='center'><B>ContentIndexState</B></td>"
Add-Content $report "</tr>"


$mbxdb = Get-MailboxDatabase | Get-MailboxDatabaseCopyStatus

$mbxdb = $mbxdb | Sort-Object Status -Descending


foreach ($db in $mbxdb)
{
	
	$dbname = $db.name
	foreach ($dbn in $dbname)
	{
		$stcopy = Get-MailboxDatabaseCopyStatus $dbn
		$srv = $stcopy.ActiveDatabaseCopy
		$mbxdbname = $stcopy.DatabaseName
		$mbxdbname1 = get-mailboxdatabase $mbxdbname
		$acpref = ($mbxdbname1 | select -ExpandProperty ActivationPreference | where { $_.key -like "$srv" }).value
	}
	
	$server = $db.Mailboxserver
	
	$status = $db.Status
	$ActiveCopy = $db.ActiveCopy
	$CopyQueuelength = $db.CopyQueuelength
	$ReplayQueueLength = $db.ReplayQueueLength
	$LastInspectedLogTime = $db.LastInspectedLogTime
	$ContentIndexState = $db.ContentIndexState
	
	$svcstatus = Get-Service -ComputerName $server msexchangesearch
	$svcsts = $svcstatus.status
	
	$result = $flow.TestMailflowResult
	$time = $Flow.MessageLatencyTime
	$remote = $Flow.IsRemoteTest
	Add-Content $report "<tr>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbname</B></td>"
	
	if ((($status -eq "Mounted") -and ($acpref -eq 1)) -or ($status -eq "Healthy"))
	{
		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$status</B></td>"
	}
	elseif ((($status -eq "Mounted") -and ($acpref -ne 1)) -or ($status -eq "Healthy"))
	{
		Add-Content $report "<td bgcolor= 'yellow' align=center>  <B>$status</B></td>"
	}
	
	else
	{
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>$status</B></td>"
		if ($ALert -eq "Yes") { Send-Email -From $from1 -To $to -subject "Open Critical - $server $dbname Unhealthy" -smtpserver $smtphost }
	}
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$ActiveCopy</B></td>"
	
	if ($CopyQueuelength -le "50")
	{
		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$CopyQueuelength</B></td>"
	}
	else
	{
		Add-Content $report "<td bgcolor= 'Yellow' align=center>  <B>$CopyQueuelength</B></td>"
	}
	
	if ($ReplayQueueLength -le "50")
	{
		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$ReplayQueueLength</B></td>"
	}
	else
	{
		Add-Content $report "<td bgcolor= 'Yellow' align=center>  <B>$ReplayQueueLength</B></td>"
	}
	
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$LastInspectedLogTime</B></td>"
	
	if (($ContentIndexState -eq "Healthy") -and ($svcsts -like "Running"))
	{
		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$ContentIndexState</B></td>"
	}
	elseif ($svcsts -notlike "Running")
	{
		Add-Content $report "<td bgcolor= 'Yellow' align=center>  <B>$svcsts</B></td>"
	}
	elseif ($ContentIndexState -ne "Healthy")
	{
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>$svcsts</B></td>"
		if ($ALert -eq "Yes") { Send-Email -From $from1 -To $to -subject "Open Critical - $server $dbname Index Unhealthy" -smtpserver $smtphost }
	}
	
	Add-Content $report "</tr>"
	
}
#################################################################################################################
##############################################DAG DB Backup Status###############################################

add-content $report  "<tr bgcolor='Lavender'>"
add-content $report  "<td colspan='7' height='25' align='center'>"
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>DAG Database Backup Status</strong></font>"
add-content $report  "</td>"
add-content $report  "</tr>"

add-content $report  "</tr>"
add-content $report  "</table>"
add-content $report  "<table width='100%'>"
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='10%' align='center'><B>Database</B></td>"
Add-Content $report  "<td width='5%' align='center'><B>BackupInProgress</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>SnapshotLastFullBackup</B></td>"
Add-Content $report  "<td width='5%' align='center'><B>SnapshotLastCopyBackup</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>LastFullBackup</B></td>"
Add-Content $report  "<td width='5%' align='center'><B>RetainDeletedItemsUntilBackup</B></td>"

Add-Content $report "</tr>"

$dbst = Get-MailboxDatabase | where{ $_.MasterType -like "DatabaseAvailabilityGroup" }

$dbst | foreach{
	$st = Get-MailboxDatabase $_ -status
	$dbname = $st.Name
	$dbbkprg = $st.BackupInProgress
	$dbsnpl = $st.SnapshotLastFullBackup
	$dbsnplc = $st.SnapshotLastCopyBackup
	$dblfb = $st.LastFullBackup
	$dbrd = $st.RetainDeletedItemsUntilBackup
	Add-Content $report "<tr>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbname</B></td>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbbkprg</B></td>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbsnpl</B></td>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbsnplc</B></td>"
	if ($dblfb -lt $hrs)
	{
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>$dblfb</B></td>"
		if ($ALert -eq "Yes") { Send-Email -From $from1 -To $to -subject "Open Critical - $dbname Backup Unhealthy" -smtpserver $smtphost }
		
	}
	else
	{
		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$dblfb</B></td>"
	}
	
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbrd</B></td>"
	Add-Content $report "</tr>"
	
	
}


##################################################################################################################
############################################## Test mail Flow For DAG ############################################


add-content $report  "<tr bgcolor='Lavender'>"
add-content $report  "<td colspan='7' height='25' align='center'>"
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Mail Flow Test Report</strong></font>"
add-content $report  "</td>"
add-content $report  "</tr>"

add-content $report  "</tr>"
add-content $report  "</table>"
add-content $report  "<table width='100%'>"
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='25%' align='center'><B>Server</B></td>"
Add-Content $report  "<td width='25%' align='center'><B>Result</B></td>"
Add-Content $report "<td width='25%' align='center'><B>Message Latency Time</B></td>"
Add-Content $report  "<td width='25%' align='center'><B>IsRemoteTest</B></td>"
Add-Content $report "</tr>"


Foreach ($server in $TestMailFlow)
{
	
	$server = $server.Name
	write-host $server
	$flow = test-mailflow $server
	if ($flow -ne $null)
	{
		
		$result = $flow.TestMailflowResult
		$time = $Flow.MessageLatencyTime
		$remote = $Flow.IsRemoteTest
		Add-Content $report "<tr>"
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$server</B></td>"
		if ($result -eq "Success")
		{
			Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B> $result</B></td>"
		}
		else
		{
			Add-Content $report "<td bgcolor= 'Red' align=center>  <B> $result</B></td>"
		}
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$time</B></td>"
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$remote</B></td>"
		
		Add-Content $report "</tr>"
	}
	
	
}
#####################################################################################################################
############################################## Get Queue For HUB Servers ############################################


add-content $report  "<tr bgcolor='Lavender'>"
add-content $report  "<td colspan='7' height='25' align='center'>"
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Mail Queue Status</strong></font>"
add-content $report  "</td>"
add-content $report  "</tr>"

add-content $report  "</tr>"
add-content $report  "</table>"
add-content $report  "<table width='100%'>"
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='25%' align='center'><B>Identity</B></td>"
Add-Content $report "<td width='25%' align='center'><B>Delivery Type</B></td>"
Add-Content $report  "<td width='25%' align='center'><B>Status</B></td>"
Add-Content $report "<td width='25%' align='center'><B>Message Count</B></td>"
Add-Content $report  "<td width='25%' align='center'><B>Next Hop Domain</B></td>"
Add-Content $report "</tr>"

$GetHub = Get-TransportServer | get-Queue

Foreach ($Queue in $GetHub)
{
	
	$Identity = $Queue.Identity
	$DeliveryType = $Queue.DeliveryType
	$Status = $Queue.Status
	$MSgCount = $Queue.Messagecount
	$NextHopDomain = $Queue.NextHopDomain
	
	
	Add-Content $report "<tr>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$DeliveryType</B></td>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$Status</B></td>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$MSgCount</B></td>"
	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$NextHopDomain</B></td>"
	
	Add-Content $report "</tr>"
	
	
}

###########################################################################################################################
######################################################### Send Mail #######################################################

Add-content $report  "</table>"
Add-Content $report "</body>"
Add-Content $report "</html>"

if ($htmlreporting -eq "Yes")
{
	$subject = "Exchange Status Check Report"
	$body = Get-Content $reportpath
	$smtp = New-Object System.Net.Mail.SmtpClient $smtphost
	$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
	$msg.isBodyhtml = $true
	$smtp.send($msg)
}
###################################################Exchange Test Complete##################################################


