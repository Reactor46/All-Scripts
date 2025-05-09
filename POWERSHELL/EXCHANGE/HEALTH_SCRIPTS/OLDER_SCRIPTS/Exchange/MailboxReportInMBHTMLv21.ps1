<#
.Synopsis
Generates HTML Report in MB.

.Description
This scripts generates a HTML report, sorted in size of the mailbox.
The script is stored at the location $ReportPath
The variable $SendEmail
$True => The Info will been send by E-mail to $MailTo
$False => No Info will been send by E-mail
The variable $Sendbody
0 => The info is NOT send in the bode of the e-mail
1 => The info is send in the body of the e-mail .
The variable $SendSeperateEmail
0 => All reports are send in 1 e-mail
1 => Each Database HTML report will be send in a seperate e-mail

.Inputs
No Inputs are necessary

.Outputs
The HTML files are the outputs

.Notes
Name: MailboxReportInMBHTMLv21.ps1
Author: Richard Voogt
Version: 2.0
Date: August 11 2013
Tested on Exchange 2007, Exchange 2010 and Exchange 2013
Written By: Richard Voogt
Website: http://www.vspbreda.nl
Twitter: http://twitter.com/rvoogt

.Change-Log
V1.0 06/01/2013 – Initial version
V1.1 06/27/2013 – hanged Subject – Initial release
V1.2 07/15/2013 – Change Exchange version selection from -eq to like (Thanks to Leon Zebregs)
V2.0 08/11/2013 – Complete redesign HTML report
Send 1 total e-mail or send an e-mail for each database
V2.1 08/12/2013 – Minor bug fixes

.Link
https://www.vspbreda.nl/nl/2013/06/create-html-mailbox-report-in-mb/

.EXAMPLE
For normal execution:
.\MailboxReportInMBHTMLv21.ps1

For testing purposes :
.\MailboxReportInMBHTMLv21.ps1 -verbose
#>

[CmdletBinding()]
param ()

<# Index
1. Define all variables
2. Function Sendmail
3. Function Alternating rows.
4. Function MailboxReportInMB
#>

#1. Define all variables

#Don’t edit between the lines
# ————————
$Date = ( get-date ).ToString(‘yyyy/MM/dd’)
$ReportDate=get-date -uformat “%Y-%m-%d-%H%M”
# ————————

# Give the appropriate E-mail settings
$MailFrom = “MailBoxSizeReport@uson.local”
$MailTo = “jbattista@optummso.com”
$MailServer = “smtp.uson.local”

# Choose your report settings
# Value 1 is active, Value 0 is not active
$SendMail = $true # Value $False is No e-mail will be send, value $true is send e-mail
$SendBody = ‘1’ # Value 0 no report in body of the e-mail, value 1 report info in body of the e-mail
$SendAttachment = ‘1’ # Value 0 no attachment report will be send, value 1 report info in a seperate attachment of the e-mail
$SendSeperateEmail = ‘0’ # Value 0 Send one E-mail, value 1 send an seperate e-mail for every Database

# Choose your misc variables
$ReportPath = “E:\Scripts\MailboxrightsInMB\reports”
$OutputFile = “Mailbox-Report-in-MB-$($ReportDate).html”
$EmailSubject = ‘Mailbox report in MB’
#$Customer = “VSPBreda”
$Subject = “$EmailSubject”
$AttachmentArray = @()

# Create report directory and Output file
New-Item $ReportPath -type directory -Force | Out-Null
Write-Verbose “Output file is created.”

#Get Exchange version
<#
$ExMajorVersionStrings = @{ 
“6.5” = @{Long=”Exchange 2003″;Short=”E2003″}
“8” = @{Long=”Exchange 2007″;Short=”E2007″}
“14” = @{Long=”Exchange 2010″;Short=”E2010″}
“15” = @{Long=”Exchange 2013″;Short=”E2013″}
}
#>
$E2010 = $false
if (Get-ExchangeServer | Where {$_.AdminDisplayVersion.Major -like “*15*”})
{
Write-Warning “Exchange 2013 or higher detected. Script will work”
$Exchangeversion="E2013"
Write-Verbose “Exchange version $Exchangeversion detected”
}
elseif (Get-ExchangeServer | Where {$_.AdminDisplayVersion -like “*14*”})
{
Write-Warning “Exchange 2010 or higher detected. Script will work”
$Exchangeversion="E2010"
Write-Verbose “Exchange version $Exchangeversion detected”
}
else
{
Write-Warning “Exchange 2007 or Lower detected. Script will NOT work !!!!!”
$Exchangeversion="E2007"
Write-Verbose “Exchange version $Exchangeversion detected”
}
Write-Verbose “Exchange version $Exchangeversion detected”

# Check if -SendMail parameter set and if so check -SendBody and -SendAttachment are set
If ($SendMail -eq $true )
{
If ( $SendBody -eq 0 -AND $SendAttachment -eq 0 )
{
cls
Write-Host -backgroundcolor Red -ForegroundColor White ‘***************************************************************************************’
Write-Host -backgroundcolor Red -ForegroundColor White ‘* *’
Write-Host -backgroundcolor Red -ForegroundColor White ‘* If -SendMail specified, you must also specify -SendBody or -SendAttachment or both *’
Write-Host -backgroundcolor Red -ForegroundColor White ‘* *’
Write-Host -backgroundcolor Red -ForegroundColor White ‘***************************************************************************************’
Write-Host ”
BREAK
} #End Check E-mail Settings
}

# Check if -SendMail parameter set and if so check -MailFrom, -MailTo and -MailServer are set
If ($SendMail)
{
If (!$MailFrom -or !$MailTo -or !$MailServer -or !$Subject)
{
cls
Write-Host -backgroundcolor Red -ForegroundColor White ‘***********************************************************************************************’
Write-Host -backgroundcolor Red -ForegroundColor White ‘* *’
Write-Host -backgroundcolor Red -ForegroundColor White ‘* If -SendMail specified, you must also specify -MailFrom, -MailTo, -MailServer and -Subject *’
Write-Host -backgroundcolor Red -ForegroundColor White ‘* *’
Write-Host -backgroundcolor Red -ForegroundColor White ‘***********************************************************************************************’
Write-Host ”
BREAK
} # End If
} #End IF $SendMailEmail

# 2. Sendmail function
# Start the function with this command:
# SendEmail -OutputFileEmail $OutputFileReport -SendMailEmail $true -mailfromEmail $MailFromReport -MailToEmail $MailToReport -MailServerEmail $MailServerReport -SubjectEmail $SubjectReport -SendBodyEmail $SendBodyReport -SendAttachmentEmail $SendAttachmentReport

Function SendEmail {
[CmdletBinding()]
Param(
[parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Output File name’)][string]$OutputFileEmail,
[parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Send Mail ($True/$False)’)][bool]$SendMailEmail,
[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail From’)][string]$MailFromEmail,
[parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail To’)]$MailToEmail,
[parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Server’)][string]$MailServerEmail,
[parameter(Position=5,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Subject’)][string]$SubjectEmail,
[parameter(Position=6,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Body’)][string]$SendBodyEmail,
[parameter(Position=7,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Attachment’)][string]$SendAttachmentEmail,
[parameter(Position=8,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Attachment Name ‘)][array]$SendAttachmentName
)

BEGIN {
} #End Begin SendEmail

Process
{
# Get the data from the HTML file and place it in ‘$body’
$body = get-content “$Reportpath\Mailbox-Report-in-MB-$($ReportDate)-$DatabaseRecord.html” | out-string

# Sending the E-mail body or attachment
write-Verbose “sendmail = $SendMail”
write-Verbose “SendBodyEmail = $SendBodyEmail”
write-Verbose “SendAttachmentEmail = $SendAttachmentEmail”
write-Verbose “MailFromEmail = $MailFromEmail”
write-Verbose “MailToEmail = $MailToEmail”
write-Verbose “SubjectEmail = $SubjectEmail”
write-Verbose “MailServerEmail = $MailServerEmail”
IF ($SendMail -eq $True)
{
IF ($SendSeperateEmail -eq 0)
{
$SendBodyEmail = ‘0’
$SubjectEmail = “$EmailSubject – $Customer”
$SendAttachmentEmail = ‘1’
}
IF ( $SendBodyEmail -eq 1 )
{
IF ( $SendBodyEmail -eq 1 -AND $SendAttachmentEmail -eq 0 )
{
Send-Mailmessage -from $MailFromEmail -to $MailToEmail -subject $SubjectEmail -smtpserver $MailServerEmail -body $body -bodyasHTML -verbose
}
ELSE
{
Send-Mailmessage -from $MailFromEmail -to $MailToEmail -subject $SubjectEmail -smtpserver $MailServerEmail -body $body -Attachments $SendAttachmentName -bodyasHTML -verbose
}
}
ELSE
{
IF ( $SendBodyEmail -eq 0 -AND $SendAttachmentEmail -eq 1 )
{
$Bodytext = ‘See the attachment’
Send-Mailmessage -from $MailFromEmail -to $MailToEmail -subject $SubjectEmail -smtpserver $MailServerEmail -body $bodytext -Attachments $SendAttachmentName -bodyasHTML -verbose
}
ELSE
{
Write-Host ”
Write-Host -backgroundcolor Green -ForegroundColor Black ‘No E-mail will be send because no Body or Attachment is set ($SendBody and $SendAttachment)’
Write-Host ”
}
}
} #End Process SendEmail
}

END {
} # End End Sendmail
} #End Function SendEmail

# 3. Function Alternating Lines.
Function Set-AlternatingLines
{
[CmdletBinding()]
Param(
[Parameter(Mandatory=$True,ValueFromPipeline=$True)]
[string]$Line,
[Parameter(Mandatory=$True)]
[string]$EvenLine,
[Parameter(Mandatory=$True)]
[string]$OddLine
)
Begin
{
$ClassName = $EvenLine
}
Process
{
If ($Line.Contains(“<tr>”))
{ $Line = $Line.Replace(“<tr>”,”<tr class=””$ClassName””>”)
If ($ClassName -eq $EvenLine)
{
$ClassName = $OddLine
}
Else
{
$ClassName = $EvenLine
}
}
Return $Line
} # End Process AlternatingLines
} # End Function Set-AlternatingLines

$Header = @”
<html>
<head>
<meta http-equiv=’Content-Type’ content=’text/html; charset=iso-8859-1′>
<title>DiskSpace Report</title>
<STYLE TYPE=”text/css”>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title> $EmailSubject </title>
</style>
</head>
<body>
<table width=’100%’>
<tr bgcolor=’#CCCCCC’>
<td colspan=’7′ height=’25’ align=’center’>
<font face=’tahoma’ color=’#003399′ size=’4′><strong>$Customer Mailbox report in MB – $date</strong></font>
</td>
</tr>
</table>
“@

# 4. Creating the Mailboxreport
# Start the function with this command:
# MailboxReportInMB -OutputFileReport $OutputFile -SendMailReport $True -MailFromReport $MailFrom -MailToReport $MailTo -MailServerReport $MailServer -SubjectReport $Subject -SendBodyReport $SendBody -SendAttachmentReport $SendAttachment

Function MailboxReportInMB
{
[CmdletBinding()]
param(
[parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Filename to write HTML report to’)][string]$OutputFileReport,
[parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Send Mail ($True/$False)’)][bool]$SendMailReport,
[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail From’)][string]$MailFromReport,
[parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail To’)]$MailToReport,
[parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Server’)][string]$MailServerReport,
[parameter(Position=5,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Subject’)][string]$SubjectReport,
[parameter(Position=6,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Body’)][string]$SendBodyReport,
[parameter(Position=7,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Attachment’)][string]$SendAttachmentReport,
[parameter(Position=8,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=’Mail Attachment Name’)][array]$SendAttachmentName
)

BEGIN
{

# Start Function writeHtmlHeader
Function writeHtmlHeader
{
param($fileName)
} # End Function WriteHtmlHeader
} # End Begin

Process {
writeHtmlHeader “$Reportpath\$OutputFileReport”
$mailboxdatabasename = Get-MailboxDatabase
foreach ($Record in $mailboxdatabasename)
{
$DatabaseRecord = $Record.name
# First write down the databasename so there is something on the screen
Write-Host ‘Starting database : ‘ $DatabaseRecord

# Get Mailboxstatistics, and convert is to HTML

# ___________
$SaveOutputFile = “Mailbox-Report-in-MB-$($ReportDate)-$DatabaseRecord.html”
$SendAttachmentName = “$Reportpath\$SaveOutputFile”
$SubjectDatabase = “$Subject – Database $DatabaseRecord”

$MailboxStatistics = Get-MailboxStatistics -database “$DatabaseRecord” | Select DisplayName, ItemCount, TotalItemSize, {$_.TotalItemSize.Value.ToMB()} | Sort-Object {$_.TotalItemSize.Value.ToMB()} -descending

$MailboxStatistics | ConvertTo-HTML -Head $Header | Set-AlternatingLines -EvenLine even -OddLine odd | Out-File “$SendAttachmentName”
$AttachmentArray = $AttachmentArray + “$SendAttachmentName”
# ___________

# First write down the databasename so there is something on the screen
Write-Host -backgroundcolor Green -ForegroundColor Black ‘Passed database : ‘ $DatabaseRecord

# Add-Content “$Reportpath\$OutputFile” $MailboxStatistics
IF ($SendMail -eq $True)
{
IF ( $SendSeperateEmail -eq 1)
{
SendEmail -OutputFileEmail “$Reportpath\$SaveOutputFile” -SendMailEmail $true -mailfromEmail $MailFromReport -MailToEmail $MailToReport -MailServerEmail $MailServerReport -SubjectEmail $SubjectDatabase -SendBodyEmail $SendBodyReport -SendAttachmentEmail $SendAttachmentReport -SendAttachmentName $SendAttachmentName
} # End IF SendSeperateEmail
} # End SendMail = True
} # End ForEach Record

IF ($SendMail -eq $True)
{
IF ( $SendSeperateEmail -eq 0)
{
Write-Verbose $MailFromReport
Write-Verbose $MailToReport
Write-Verbose $MailServerReport
Write-Verbose $SubjectDatabase
Write-Verbose $SendBodyReport
Write-Verbose $SendAttachmentReport

# Only for testing purposes : (Remove the text option below)
<#
ForEach ( $records in $AttachmentArray) {
Write-Host “Items = $records”
}
#>

$Bodytext = ‘See the attachment’
SendEmail -OutputFileEmail “$Reportpath\$SaveOutputFile” -SendMailEmail $true -mailfromEmail $MailFromReport -MailToEmail $MailToReport -MailServerEmail $MailServerReport -SubjectEmail $SubjectDatabase -SendBodyEmail $Bodytext -SendAttachmentEmail $SendAttachmentReport -SendAttachmentName $AttachmentArray
} # End IF $SendSeperateEmail
} # End SendMail = True
} # End Process

END {
Write-Host ”
Write-Host -backgroundcolor Green -ForegroundColor Black ‘Files saved at : ‘ $Reportpath
Write-Host ”
} # End End
} #End Function

# Start the script
write-Verbose “Sendmail = $SendMail”
IF ($SendMail -eq $True)
{
MailboxReportInMB -OutputFileReport $OutputFile -SendMailReport $True -MailFromReport $MailFrom -MailToReport $MailTo -MailServerReport $MailServer -SubjectReport $Subject -SendBodyReport $SendBody -SendAttachmentReport $SendAttachment -SendAttachmentName $SendAttachmentName
}