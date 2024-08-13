 
 <#Monitor-GPOChanges.ps1
 
Purpose : Audit GPO events. Heavily adapted from https://gallery.technet.microsoft.com/scriptcenter/Audit-Active-Directory-3f4a17dd 
Exports a rather crude HTML report. Intended to run as a scheduled task, so use appropriate permissions

The purpose of this project is to audit all Active Directory changes regarding GPO management and display these changes
The powershell script extract the following events :

5136         > A directory service object was modified
5137         > A directory service object was created
5138         > A directory service object was undeleted
5139         > A directory service object was moved
5141         > A directory service object was deleted
 
Pre-requisites :
Tested on Domain controller running Windows 2008/2008R2
Powershell v3
Active Directory Module on running computer
Group Policy Management feature installed (servers) -OR- RSAT (workstations)

References:
https://gallery.technet.microsoft.com/scriptcenter/Audit-Active-Directory-3f4a17dd 
https://blogs.technet.microsoft.com/ashleymcglone/2013/08/28/powershell-get-winevent-xml-madness-getting-details-from-event-logs/
https://blogs.technet.microsoft.com/ashleymcglone/2015/08/31/forensics-automating-active-directory-account-lockout-search-with-powershell-an-example-of-deep-xml-filtering-of-event-logs-across-multiple-servers-in-parallel/

NOTE: There are references to making an object array with the report data, which I commented out. 
Leaving this as-is for now, but can use the object array for making CSV output or fanicer HTML output.

#> 
####################
# USER VARIABLES   #
####################

$DaystoKeepEventData = 60 # Please modify to your liking
$ADDomain = "Contoso.corp" # Required
$PlaceToCopyFullReport = "C:\Scripts\Repository\jbattista\Web\Reports\" # if this is blank, it will be ignored
$global:smtpServer = "mailgateway.Contoso.corp"
$global:FromAddress = "GPOAudit@creditone.com"
$global:DestMailboxes = "john.battista@creditone.com" #@("<you@contoso.com>","<someoneelse@contoso.com>") 

#-----------------------------------------------------------------------------------#
#                                                                                   # 
#	DO NOT EDIT ANY OF THE FOLLOWING CODE - EDITS ARE TO BE MADE ABOVE HERE ONLY!!  #
#                                                                                   #
#-----------------------------------------------------------------------------------#

####################
# LOAD MODULES     #
####################

Import-Module -Name activedirectory
#Import-Module -Name grouppolicy



######################
# SCRIPT VARIABLES   #
######################

$HTMLoutputpath = 'C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\GPOMonitor\HTML\'
$HTMLHEADERoutputfilename = 'C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\HTML\HTMLHEADERGPOEvents.htm'
$HTMLTABLEoutputfilename = "C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\HTML\" +"{0:yyyy-MM-dd-hh-mm}" -f (get-date)+"GPOEvents.htm"
$HTMLFullReportFilename = 'C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\HTML\FullReportGPOEvents.htm'
$CSVoutputpath = 'C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\CSV\'
$CSVHEADERoutputfilename = 'C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\CSV\CSVHEADERGPOEvents.CSV'
$CSVTABLEoutputfilename = "C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\CSV\" +"{0:yyyy-MM-dd-hh-mm}" -f (get-date)+"GPOEvents.htm"
$CSVFullReportFilename = 'C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\CSV\FullReportGPOEvents.CSV'
$GPOListingFilename = 'C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\HTML\GPOHistoricalRoster.CSV'
$TodaysDate = "{0:yyyy-MM-dd hh:mm tt}" -f (get-date)
$GPOHistoryOutputPath = 'C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\GPOHistory\' 
$GPOHistoryListing = "C:\Scripts\Repository\jbattista\Web\Reports\GPOMonitor\GPOHistory\GPOHistory.TXT"

#######################
# SET SCRIPT PATH     #
#######################

$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Write-host "My directory is $dir"
# temporarily change to the correct folder
Set-Location $dir

#######################
# FIND LAST RUNTIME   #
#######################
Write-Host "Finding the last runtime..." -ForegroundColor Green
IF ($HTMLFullReportFilename ){ #the current one out there is the old one.

$LastRuntime = Get-Content ($HTMLoutputpath + "\") | sort LastWriteTime | select -last 1 | select LastWriteTime
write-host "Last runtime is: "$LastRuntime 
write-host "Last runtime type: "$LastRuntime.GetType()
write-host "Get-Date is: " (get-date)
write-host "Get-Date type: " (get-date).GetType()
$LastRuntime  = $LastRuntime -replace '=',"." 
                            $LastRuntime  = $LastRuntime -replace '}',"." 
                            $LastRuntime = ($LastRuntime.split('.'))
                            $LastRuntime = $LastRuntime[1]
                            $LastRuntime = [datetime]$LastRuntime 
write-host "Last runtime is: "$LastRuntime 
write-host "Last runtime type: "$LastRuntime.GetType()
                      

$TimeDifference = (New-TimeSpan -Start $LastRuntime -End (get-date)).TotalMilliseconds
write-host "Time difference is: " $TimeDifference "ms"
}
####################
#   MAIL FUNCTION   #
#####################

Function EmailReport (){
    
    Param
    (
    [String]$CurrentEmailAddress,
	[String]$EmailBody
    )
    $msg = new-object Net.Mail.MailMessage
    $smtp = new-object Net.Mail.SmtpClient($global:smtpServer)
	$msg.IsBodyHTML = $true
	$msg.Subject = "Group Policy Change " 
	$msg.Body = $EmailBody
	$msg.To.Add($CurrentEmailAddress)
    $msg.From = $global:FromAddress
    $smtp.Send($msg)
    }


###############################
# DEFINE THE EVENT XML QUERY  #
###############################

$XMLQuery='<QueryList>
            <Query Id="0" Path="Security">
                <Select Path="Security">
                    *[System[(EventID=5136 or EventID=5137 or EventID=5138 or EventID=5139 or EventID=5141) and TimeCreated[timediff(@SystemTime) &lt;= ' + $TimeDifference + ' ]]] 
                and 
                    *[EventData[Data[@Name="ObjectClass"] and (Data="groupPolicyContainer")]]
                </Select>
            </Query>
           </QueryList>'

#######################
# MAKE OUTPUT FOLDERS #
#######################

if ((Test-Path -LiteralPath $HTMLoutputpath ) -eq $true) {
write-host 'HTML output dir found' -ForegroundColor green
}
else
{
write-Host "No HTML output dir found - making a new one" -ForegroundColor green
New-Item -ItemType directory -Path $HTMLoutputpath
}

if ((Test-Path -LiteralPath $CSVoutputpath ) -eq $true) {
write-host 'CSV output dir found' -ForegroundColor green
}
else
{
write-Host "No CSV output dir found - making a new one" -ForegroundColor green
New-Item -ItemType directory -Path $CSVoutputpath
}

if ((Test-Path -LiteralPath $GPOHistoryOutputPath  ) -eq $true) {
write-host 'GPO History Listing output dir found' -ForegroundColor green
}
else
{
write-Host "No GPO History Listing output dir found - making a new one" -ForegroundColor green
New-Item -ItemType directory -Path $GPOHistoryOutputPath 
}


###################
# FILE MANAGEMENT #
###################

$limit = (get-date).AddDays(-($DaystoKeepEventData))
write-host  "Deleting Files older than " $limit "..." -ForegroundColor Green
Get-ChildItem -path $HTMLoutputpath -Recurse -force | Where-object {!$_.psiscontainer -and $_.CreationTime -lt $limit } | Remove-item -force 
if ((Test-Path $HTMLFullReportFilename) -eq $true) {Remove-item $HTMLFullReportFilename -force } #remove the old report
if ((Test-Path $HTMLHEADERoutputfilename) -eq $true) {Remove-item $HTMLHEADERoutputfilename -force } #remove the old report

###################
# FIRST RUN       #
###################

if ((Test-Path $GPOHistoryListing ) -eq $false) {
$FirstRun = "True" #Set the marker for later
$CurrentGPOListing = @(Get-GPO -all)
foreach ($CurrentGPO in $CurrentGPOListing)
    {
    $CurrentGPOListing = "<TR><TD>"+$CurrentGPO.DisplayName +"</TD><TD>"+$CurrentGPO.ID+"</TD></TR>" + ([Environment]::NewLine)
    Add-Content $GPOHistoryListing $CurrentGPOListing
    }
} 
else
{

###################
# CHECK EVENTS    #
###################

Write-Host "Getting List of DCs..." -ForegroundColor Yellow

$DCServers = $DomainControllers = Get-ADDomainController -Filter *

Foreach ($DCServer in $DomainControllers)
{
Write-Host "Getting events from " $DCServer -ForegroundColor Yellow

 Try {      $Events = Get-WinEvent -ComputerName $DCServer -FilterXml $XMLQuery -ErrorAction Stop
            Write-Host "Events found - now I'm parsing their data" -ForegroundColor Green

            ForEach ($Event in $Events) {
                
                # Convert the event to XML
                $eventXML = [xml]$Event.ToXml()
                
                # Iterate through each one of the XML message properties
                For ($i=0; $i -lt $eventXML.Event.EventData.Data.Count; $i++) {
                    # Append these as object properties
                    Add-Member -InputObject $Event -MemberType NoteProperty -Force `
                        -Name  $eventXML.Event.EventData.Data[$i].name `
                        -Value $eventXML.Event.EventData.Data[$i].'#text'
                }
                $date = "{0:yyyy-MM-dd-hh:mm tt}" -f ($event.TimeCreated)
                $id = $event.id.ToString()
                               
                # We need to convert the GPO DN to a Friendly Name
                $GPODN = $Event.objectDN
                $GPODN  = $GPODN -replace '}',"{" 
                $GPOGUIDtext = ($GPODN.split('{')) 
                $GPOGUID = [GUID]($GPOGUIDtext[1]) 
                $GetGPOName = Get-GPO -Guid $GPOGUID -ErrorAction SilentlyContinue
                if ($GetGPOName){
                $GPOName= $GetGPOName | Select DisplayName
                    # There's a GPO there, let's find its name
                            $GPOName  = $GPOName -replace '=',"{" 
                            $GPOName  = $GPOName -replace '}',"{" 
                            $GPONameArray = ($GPOName.split('{'))
                            $GPOFriendlyName = $GPONameArray[2]
                            Write-Host "Got GPO Name - and info" -ForegroundColor DarkYellow
                    }
                    else{     #If this fails, it is probably because the GPO was deleted
                    Write-Host "GPO was deleted" -ForegroundColor DarkYellow
                    $GPOFriendlyName = "[NAMENOTFOUND] - You can find in the history table below"
                    }
                $CulpritUserID = $Event.SubjectUserName
                $CulpritUserGivenName = (Get-ADUser -Identity $CulpritUserID -properties GivenName).givenname
                $CulpritUserSurname = (Get-ADUser -Identity $CulpritUserID -properties Surname).surname
                $CulpritUserFullName = $CulpritUserGivenName + " " + $CulpritUserSurname 
                
                
                switch ($id){ #Now let's start reporting this stuff...
                    ("5136"){$Eventmeaning = "MODIFIED"}
                    ("5137"){$Eventmeaning = "CREATED"
                            If ($FirstRun -ne "True") 
                                {
                                $NewGPOInfo = "<TR><TD>"+$GPOFriendlyName +"</TD><TD>"+$GPOGUIDtext+"</TD></TR>"
                                Add-Content $GPOHistoryListing $NewGPOInfo 
                                }
                            }
                    ("5138"){$Eventmeaning = "UNDELETED"}
                    ("5139"){$Eventmeaning = "MOVED"}
                    ("5141"){$Eventmeaning = "DELETED"}
                }
                $HTMLOutputLine = "<TR><TD> "+ $date +" </TD><TD><STRONG> " + $CulpritUserFullName +" </STRONG></TD><TD> " +  $GPOFriendlyName  +" </TD><TD><STRONG> " + $Eventmeaning +" </STRONG></TD><TD> " + $GPOGUID + " </TD></TR>" 
                $HTMLNewTableEntry += $HTMLOutputLine
				$MailBody = "When: "+ $date +"<BR>Who: " + $CulpritUserFullName +" <BR>GPO: " +  $GPOFriendlyName  +" <BR>How: " + $Eventmeaning +" <BR>GUID: " + $GPOGUID + " <P><A HREF=""http://webconnections/Reports/FullReportGPOEvents.htm"">CLICK HERE FOR MORE INFO</A>" 
				Foreach ($mailaddress in $global:DestMailboxes){EmailReport $mailaddress $MailBody}
				
            }
        
        }
        Catch {
            If ($_.Exception -like "*No events were found that match the specified selection criteria*") {
                Write-Host "No events found since last time. I shall contemplate my navel..." -ForegroundColor Green
            } Else {$_}
        }
}
}

###############
# HTML OUTPUT #
###############

#Sort all the output files by date, then stick them into the full report
ForEach ($sourcefile In $(Get-ChildItem $HTMLoutputpath | Sort-Object -Property CreationTime -Descending))
{
    write-host "Adding older report " $sourcefile " to the full report..." -ForegroundColor Yellow
    $HTMLOldTable += Get-Content ($HTMLoutputpath + "\" + $sourcefile) -Raw 
}

write-host "Making GPO History Listing..." -ForegroundColor Yellow
$GPOHistoryTable = Get-Content $GPOHistoryListing | sort #| get-unique # - that probably is not necessary, but could be added later

# NOW make the full report
$HTMLHeader1 = '<HTML><BODY><font face="verdana"><H3>GPO Changes for the last ' + $DaystoKeepEventData + ' days - last modified: ' + $TodaysDate + '</H3>'
$HTMLHeader2 = '<TABLE><TR><strong><TD>DateTime</TD><TD>User Name</TD><TD>Policy Name</TD><TD>Change Type</TD><TD>GPO GUID</TD></strong></TR>'
$HTMLFullTable = $HTMLOldTable  #Tack on the current report to the HTML table
$HTMLTableEnd = '</TABLE>'
$HTMLFooter = '</BODY></HTML>'
$HTMLHistoryListingTableStart = '<H3>GPO Historical Listing</H3><TABLE><TR><strong><TD>Policy Name</TD><TD>ID</TD></strong></TR>'
Add-Content $HTMLFullReportFilename $HTMLHeader1
Add-Content $HTMLFullReportFilename $HTMLHeader2
Add-Content $HTMLFullReportFilename $HTMLNewTableEntry 
Add-Content $HTMLFullReportFilename $HTMLOldTable 
Add-Content $HTMLFullReportFilename $HTMLTableEnd 
Add-Content $HTMLFullReportFilename $HTMLHistoryListingTableStart
Add-Content $HTMLFullReportFilename $GPOHistoryTable
Add-Content $HTMLFullReportFilename $HTMLTableEnd 
Add-Content $HTMLFullReportFilename $HTMLFooter

Write-Host "Finished writing full report to: " $HTMLFullReportFilename -ForegroundColor Green

Write-Host "Writing new output to: " $HTMLFullReportFilename -ForegroundColor Yellow
Add-Content $HTMLTABLEoutputfilename $HTMLNewTableEntry 
Write-Host "Finished writing output to: " $HTMLTABLEoutputfilename -ForegroundColor Green

#IF you have a secondary place ($PlaceToCopyFullReport) where you want the report to go, copy the report there

if (Test-Path $PlaceToCopyFullReport) {copy-item -path $HTMLFullReportFilename -Destination $PlaceToCopyFullReport -Force}
else {Write-Host "No secondary path is specified for the HTML report"}

################
# CLEANUP VARS #
################
if ($DaystoKeepEventData) { try { Remove-Variable -Name DaystoKeepEventData -Scope Global -Force } catch { } }
if ($DCServer) { try { Remove-Variable -Name DCServer -Scope Global -Force } catch { } }
if ($Interval) { try { Remove-Variable -Name Interval -Scope Global -Force } catch { } }
if ($PlaceToCopyFullReport) { try { Remove-Variable -Name PlaceToCopyFullReport -Scope Global -Force } catch { } }
if ($HTMLoutputpath) { try { Remove-Variable -Name HTMLoutputpath -Scope Global -Force } catch { } }
if ($HTMLHEADERoutputfilename) { try { Remove-Variable -Name HTMLHEADERoutputfilename -Scope Global -Force } catch { } }
if ($HTMLTABLEoutputfilename) { try { Remove-Variable -Name HTMLTABLEoutputfilename -Scope Global -Force } catch { } }
if ($HTMLFullReportFilename) { try { Remove-Variable -Name HTMLFullReportFilename -Scope Global -Force } catch { } }
if ($TodaysDate) { try { Remove-Variable -Name TodaysDate -Scope Global -Force } catch { } }
if ($XMLQuery) { try { Remove-Variable -Name XMLQuery -Scope Global -Force } catch { } }
if ($limit) { try { Remove-Variable -Name limit -Scope Global -Force } catch { } }
if ($Events) { try { Remove-Variable -Name Events -Scope Global -Force } catch { } }
if ($eventXML) { try { Remove-Variable -Name eventXML -Scope Global -Force } catch { } }
if ($date) { try { Remove-Variable -Name date -Scope Global -Force } catch { } }
if ($GPODN) { try { Remove-Variable -Name GPODN -Scope Global -Force } catch { } }
if ($GPOGUIDtext) { try { Remove-Variable -Name GPOGUIDtext -Scope Global -Force } catch { } }
if ($GPOGUID) { try { Remove-Variable -Name GPOGUID -Scope Global -Force } catch { } }
if ($GetGPOName) { try { Remove-Variable -Name GetGPOName -Scope Global -Force } catch { } }
if ($GPOName) { try { Remove-Variable -Name GPOName -Scope Global -Force } catch { } }
if ($GPONameArray) { try { Remove-Variable -Name GPONameArray -Scope Global -Force } catch { } }
if ($GPOFriendlyName) { try { Remove-Variable -Name GPOFriendlyName -Scope Global -Force } catch { } }
if ($id) { try { Remove-Variable -Name id -Scope Global -Force } catch { } }
if ($Eventmeaning) { try { Remove-Variable -Name Eventmeaning -Scope Global -Force } catch { } }
if ($HTMLOutputLine) { try { Remove-Variable -Name HTMLOutputLine -Scope Global -Force } catch { } }
if ($HTMLNewTableEntry) { try { Remove-Variable -Name HTMLNewTableEntry -Scope Global -Force } catch { } }
if ($HTMLHeader1) { try { Remove-Variable -Name HTMLHeader1 -Scope Global -Force } catch { } }
if ($HTMLHeader2) { try { Remove-Variable -Name HTMLHeader2 -Scope Global -Force } catch { } }
if ($HTMLFooter) { try { Remove-Variable -Name HTMLFooter -Scope Global -Force } catch { } }
if ($HTMLFullTable) { try { Remove-Variable -Name HTMLFullTable -Scope Global -Force } catch { } }
if ($PlaceToCopyFullReport) { try { Remove-Variable -Name PlaceToCopyFullReport -Scope Global -Force } catch { } }
if ($LastRuntime) { try { Remove-Variable -Name LastRuntime -Scope Global -Force } catch { } }
if ($TimeDifference) { try { Remove-Variable -Name TimeDifference -Scope Global -Force } catch { } }
if ($CulpritUserFullName) { try { Remove-Variable -Name CulpritUserFullName -Scope Global -Force } catch { } }
if ($HTMLOldTable) { try { Remove-Variable -Name HTMLOldTable -Scope Global -Force } catch { } }
if ($GPOHistoryListing) { try { Remove-Variable -Name GPOHistoryListing -Scope Global -Force } catch { } }
if ($GPOHistoryOutputPath) { try { Remove-Variable -Name GPOHistoryOutputPath -Scope Global -Force } catch { } }
if ($HTMLGPOHistory) { try { Remove-Variable -Name HTMLGPOHistory -Scope Global -Force } catch { } }
if ($HTMLTableEnd ) { try { Remove-Variable -Name HTMLTableEnd -Scope Global -Force } catch { } }
if ($GPOHistoryTable) { try { Remove-Variable -Name GPOHistoryTable -Scope Global -Force } catch { } }
if ($HTMLHistoryListingTableStart) { try { Remove-Variable -Name HTMLHistoryListingTableStart -Scope Global -Force } catch { } }


#######
# END #
#######