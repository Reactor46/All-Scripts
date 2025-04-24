####################################################################################
# SharePoint + Eventviewer logs mailer from localhost and remotely logs via e-mail #
# by Michał Demczyszak @ Newind 2017 #
####################################################################################
 
#Set up SharePoint Shell Snap in for Powershell
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
 
#Set up SharePoint CA
$caWebApp = (Get-SPWebApplication -IncludeCentralAdministration) | ? { $_.IsAdministrationWebApplication -eq $true }
$caWeb = Get-SPWeb -Identity $caWebApp.Url 
 
#Set up method of checking event viewer logs for localhost app server and remotely sql server. Checking peroids od reading event viewer logs.
if ( (get-date).dayofweek -eq "Monday" ) 
{ 
$after = [DateTime]::Today.AddDays(-3).AddHours(7)
$before = [DateTime]::Today.AddHours(9) 
#$today = Get-Date -DisplayHint DateTime
$event = get-eventlog -LogName System -entrytype Error -after $after -before $before
$machine = "YOURSQLDATABASESERVERNAME"
$eventDB = get-eventlog -ComputerName $Machine -LogName System -entrytype Error -after $after -before $before
}
 
else
{
$after = [DateTime]::Today.AddDays(-1).AddHours(7)
$before = [DateTime]::Today.AddHours(9) 
#$today = Get-Date -DisplayHint DateTime
$event = get-eventlog -LogName System -entrytype Error -after $after -before $before
$machine = "YOURSQLDATABASESERVERNAME"
$eventDB = get-eventlog -ComputerName $Machine -LogName System -entrytype Error -after $after -before $before
}
 
#Set up email from, email to and SMTP Server Addresses
$toAddress = "ADMIN1@DOMAIN.COM, ADMIN2@DOMAIN.COM"
$fromAddress = $caWebApp.OutboundMailReplyToAddress
$serverAddress = $caWebApp.OutboundMailServiceInstance.Server.Address
 
#Get Health Analyzer list on Central Admin site
$healthList = $caWeb.GetList("\Lists\HealthReports")
$displayFormUrl = $caWeb.Url + ($healthList.Forms | where { $_.Type -eq "PAGE_DISPLAYFORM" }).ServerRelativeUrl
#Set up query to Health Analyzer list on Central Admin site
$queryString = "4 - Powodzenie"
$query = New-Object Microsoft.SharePoint.SPQuery
$query.Query = $queryString
$items = $healthList.GetItems($query)
 
$msgTitle = "YOUR RAPORT EMAIL TITLE NAME" + (Get-Date)
#HTML head
$head = "<style type=`"text/css`">.tableStyle { border: 1px solid #000000; }</style>"
$head = $head + "<Title>$msgTitle $Outputreport</Title>"
 
#Create HTML body by walking through each item and adding it to a table
$body = "<H2>$msgTitle</H2>
<table cellspacing=`"0`" class=`"tableStyle`" style=`"width: 100%`">"
foreach ($item in $items)
{
 $itemUrl = $displayFormUrl + "?id=" + $item.ID
 [array]$itemValues = @($item["Ważność"], $item["Kategoria"], $item["Wyjaśnienie"], $item["Modified"]) 
 $body = $body + "
<tr>"
 $body = $body + "
<td class=`"tableStyle`"><a href=`"" + $itemUrl + "`">" + $item.Title + "</a></td>
"
 $itemValues | ForEach-Object {
 $body = $body + "
<td class=`"tableStyle`">$_</td>
"
 }
 $body = $body + "</tr>
"
}
#Create HTML body for event viewer logs from Application Server. 
$rows = ""
foreach ($eventEntry in $event){
$rows += @"
<tr>
<td style="text-align: center; padding: 5px;">$($eventEntry.TimeGenerated)</td>
<td style="text-align: center; padding: 5px;">$($eventEntry.EntryType)</td>
<td style="text-align: center; padding: 5px;">$($eventEntry.Source)</td>
<td style="padding: 5px;">$($eventEntry.Message)</td>
</tr>
"@ 
}
$email = @"
<table style="width:100%;border">
<tr>
<th style="text-align: center; padding: 5px;">Time</th>
<th style="text-align: center; padding: 5px;">Type</th>
<th style="text-align: center; padding: 5px;">Source</th>
<th style="text-align: center; padding: 5px;">SERWER APLIKACYJNY</th>
</tr>
$rows</table>
"@
 
#Create HTML body for event viewer logs from SQL Database Server.
$rowsDB = ""
foreach ($eventEntryDB in $eventDB){
$rowsDB += @"
<tr>
<td style="text-align: center; padding: 5px;">$($eventEntryDB.TimeGenerated)</td>
<td style="text-align: center; padding: 5px;">$($eventEntryDB.EntryType)</td>
<td style="text-align: center; padding: 5px;">$($eventEntryDB.Source)</td>
<td style="padding: 5px;">$($eventEntryDB.Message)</td>
</tr>
"@ 
}
$emailDB = @"
<table style="width:100%;border">
<tr>
<th style="text-align: center; padding: 5px;">Time</th>
<th style="text-align: center; padding: 5px;">Type</th>
<th style="text-align: center; padding: 5px;">Source</th>
<th style="text-align: center; padding: 5px;">SERWER BAZODANOWY</th>
</tr>
$rowsDB</table>
"@
 
#Create HMTL body for e-mail.
$body = $body + "</table>
"
$body = $body + $email + $emailDB
 
#Create message body using the ConvertTo-Html PowerShell cmdlet
$msgBody = ConvertTo-Html -Head $head -Body $body
 
#Create e-mail message object using System.Net.Mail class.
$msg = New-Object System.Net.Mail.MailMessage($fromAddress, $toAddress, $msgTitle, $msgBody)
$msg.IsBodyHtml = $true
 
#Send message.
$smtpClient = New-Object System.Net.Mail.SmtpClient($serverAddress)
$smtpClient.Send($msg)
$caWeb.Dispose()