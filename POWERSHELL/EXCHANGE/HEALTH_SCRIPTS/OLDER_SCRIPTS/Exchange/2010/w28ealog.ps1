[System.Reflection.Assembly]::LoadWithPartialName("System.Core")

$ServerName = "servername"
$DateFrom = [System.DateTime]::Now.addDays(-2)
$DateTo = [System.DateTime]::Now.addDays(-3)
$DateFromTS = New-TimeSpan -Start $DateFrom -End ([System.DateTime]::Now)
$DateToTS = New-TimeSpan -Start $DateTo -End ([System.DateTime]::Now)
$eacombCollection = @()

$elLogQuery = "<QueryList><Query Id=`"0`" Path=`"Security`"><Select Path=`"Exchange Auditing`">*[System[(EventID=10100) and TimeCreated[timediff(@SystemTime) &lt;=" + $DateToTS.TotalMilliseconds +"] and TimeCreated[timediff(@SystemTime) &gt;= " + $DateFromTS.TotalMilliseconds +"]]]</Select></Query></QueryList>"
$eqEventLogQuery = new-object System.Diagnostics.Eventing.Reader.EventLogQuery("Exchange Auditing", [System.Diagnostics.Eventing.Reader.PathType]::LogName, $elLogQuery);
$eqEventLogQuery.Session = new-object System.Diagnostics.Eventing.Reader.EventLogSession($ServerName);
$lrEventLogReader  = new-object System.Diagnostics.Eventing.Reader.EventLogReader($eqEventLogQuery)

for($eventInstance = $lrEventLogReader.ReadEvent();$eventInstance -ne $null; $eventInstance = $lrEventLogReader.ReadEvent()){
	[System.Diagnostics.Eventing.Reader.EventLogRecord]$erEventRecord =  [System.Diagnostics.Eventing.Reader.EventLogRecord]$eventInstance
	if($erEventRecord.Properties[5].Value -match "<NULL>" -eq $false){
		$exAuditObject = "" | select RecordID,TimeCreated,FolderPath,FolderName,Mailbox,AccessingUser,MailboxLegacyExchangeDN,AccessingUserLegacyExchangeDN,MachineName,Address,ProcessName,ApplicationId
		$exAuditObject.RecordID = $erEventRecord.RecordID
		$exAuditObject.TimeCreated = $erEventRecord.TimeCreated
		$exAuditObject.FolderPath = $erEventRecord.Properties[0].Value.ToString()
		$exAuditObject.FolderName = $erEventRecord.Properties[1].Value.ToString()
		$exAuditObject.Mailbox = $erEventRecord.Properties[2].Value.ToString()
		$exAuditObject.AccessingUser = $erEventRecord.Properties[3].Value.ToString()
		$exAuditObject.AccessingUserLegacyExchangeDN = $erEventRecord.Properties[4].Value.ToString()
		$exAuditObject.MailboxLegacyExchangeDN = $erEventRecord.Properties[5].Value.ToString()
		$exAuditObject.MachineName = $erEventRecord.Properties[8].Value.ToString()
		$exAuditObject.Address = $erEventRecord.Properties[9].Value.ToString()
		$exAuditObject.ProcessName = $erEventRecord.Properties[10].Value.ToString()
		$exAuditObject.ApplicationId = $erEventRecord.Properties[12].Value.ToString()
		$eacombCollection +=$exAuditObject
	}
}

$eacombCollection | export-csv –encoding "unicode" -noTypeInformation "c:\temp\exauditOutput.csv"