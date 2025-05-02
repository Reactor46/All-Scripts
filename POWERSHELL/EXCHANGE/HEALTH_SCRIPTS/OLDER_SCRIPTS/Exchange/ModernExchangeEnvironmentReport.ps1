<#

.Requires -version 2 - Runs in Exchange Management Shell

.SYNOPSIS
.\ModernExchangeEnvironmentReport - 

It displays Complete Exchange Environment Information in a modern HTML. It covers only Exchange 2010 or later.

Sample Report can be seen at - http://www.careexchange.in/wp-content/uploads/2015/09/ModernExchangeEnvironmentReport.htm

.Author
Written By: Satheshwaran Manoharan

Change Log
V1.0, 28/09/2015 - Initial version

Change Log
V1.1, 02/07/2016 - Added "Enter the Distribution Group name with Wild Card"

Change Log
v1.2 11/10/2016 - Added Date to File Name

#>

#Add Exchange Server snapin if not already loaded

if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
}
Write-Progress -Activity "Date" -status "Collecting Dates info"
$Date = Get-date -ErrorAction SilentlyContinue

# ----- Settings ----

#Should the Script Open the HTML File locally on Finishing the script - Say Yes - if you wish to

$Openhtmllocally = "No"

Write-Progress -Activity "Email Settings" -status "Storing Email Settings"
## ----- Email ----Fill in with your details  

$EmailTo = "administrator@dubai.com"
$EmailFrom = "administrator@dubai.com"
$EmailSubject = "Modern Exchange Environement Report $Date"
# Creating Anonymous Relay - http://www.careexchange.in/how-to-configure-a-relay-connector-for-exchange-server-2013/
$SmtpServer = "10.128.57.55"
$Date = (Get-Date -format "MM-dd-yyyy")
$Attachment = "C:\scripts\ModernExchangeEnvironmentReport($date).htm"

# ----- Settings ----

#Collecting Environment Information 
Write-Progress -Activity "Exchange Servers" -status "Collecting Exchange Servers info"
$ExchangeServers = Get-ExchangeServer -ErrorAction SilentlyContinue

Write-Progress -Activity "Mailboxes" -status "Collecting Collecting Mailboxes info"
$mailboxes = Get-mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue

Write-Progress -Activity "Databases" -status "Collecting Databases info"
$Databases = Get-mailboxdatabase -Status -ErrorAction SilentlyContinue

Write-Progress -Activity "DAG" -status "Collecting DAG info"
$DAGS = Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue

Write-Progress -Activity "DistributionGroups" -status "Collecting DistributionGroup info"
$DistributionGroups = Get-DistributionGroup -ResultSize Unlimited -ErrorAction SilentlyContinue

Write-Progress -Activity "DynamicDistributionGroups" -status "Collecting DynamicDistributionGroup info"
$DynamicGroups =  Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction SilentlyContinue

Write-Progress -Activity "MailContacts" -status "Collecting MailContacts info"
$MailContacts = Get-MailContact -ResultSize Unlimited -ErrorAction SilentlyContinue

Write-Progress -Activity "Public Folder" -status "Collecting Public Folder Mailboxes info"
$PFmailboxes = Get-mailbox -PublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue

Write-Progress -Activity "SendConnectors" -status "Collecting SendConnectors info"
$SendConnectors = Get-SendConnector -ErrorAction SilentlyContinue

Write-Progress -Activity "Accepted Domains" -status "Collecting Accepted Domains info"
$AcceptedDomains = Get-AcceptedDomain -ErrorAction SilentlyContinue

Write-Progress -Activity "OrgAdmins" -status "Collecting OrgAdmins info"
$OrgAdmins = Get-RoleGroupMember "Organization Management" -ErrorAction SilentlyContinue

#Applying Initial CSS For the HTML
Write-Progress -Activity "ModernExchangeEnvironmentReport" -status "Applying CSS"
$head = @"
<Title>Volume Report</Title>
<Style>
.CSSTableGenerator {
	margin:0px;padding:0px;
	width:100%;
	box-shadow: 10px 10px 5px #888888;
	border:1px solid #7aa3c1;
	
	-moz-border-radius-bottomleft:0px;
	-webkit-border-bottom-left-radius:0px;
	border-bottom-left-radius:0px;
	
	-moz-border-radius-bottomright:0px;
	-webkit-border-bottom-right-radius:0px;
	border-bottom-right-radius:0px;
	
	-moz-border-radius-topright:0px;
	-webkit-border-top-right-radius:0px;
	border-top-right-radius:0px;
	
	-moz-border-radius-topleft:0px;
	-webkit-border-top-left-radius:0px;
	border-top-left-radius:0px;
}.CSSTableGenerator table{
    border-collapse: collapse;
    border-spacing: 0;
	width:100%;
	height:100%;
	margin:0px;padding:0px;
}.CSSTableGenerator tr:last-child td:last-child {
	-moz-border-radius-bottomright:0px;
	-webkit-border-bottom-right-radius:0px;
	border-bottom-right-radius:0px;
}
.CSSTableGenerator table tr:first-child td:first-child {
	-moz-border-radius-topleft:0px;
	-webkit-border-top-left-radius:0px;
	border-top-left-radius:0px;
}
.CSSTableGenerator table tr:first-child td:last-child {
	-moz-border-radius-topright:0px;
	-webkit-border-top-right-radius:0px;
	border-top-right-radius:0px;
}.CSSTableGenerator tr:last-child td:first-child{
	-moz-border-radius-bottomleft:0px;
	-webkit-border-bottom-left-radius:0px;
	border-bottom-left-radius:0px;
}.CSSTableGenerator tr:hover td{
	
}
.CSSTableGenerator tr:nth-child(odd){ background-color:#ffffff; }
.CSSTableGenerator tr:nth-child(even){ background-color:#ffffff; }
.CSSTableGenerator td{
	vertical-align:middle;	
	border:1px solid #7aa3c1;
	border-width:0px 1px 1px 0px;
	text-align:center;
	padding:8px;
	font-size:10px;
	font-family:Arial;
	font-weight:normal;
	color:#000000;
}.CSSTableGenerator tr:last-child td{
	border-width:0px 1px 0px 0px;
}.CSSTableGenerator tr td:last-child{
	border-width:0px 0px 1px 0px;
}.CSSTableGenerator tr:last-child td:last-child{
	border-width:0px 0px 0px 0px;
}
.CSSTableGenerator tr:first-child td{
	background:-o-linear-gradient(bottom, #0072c6 5%, #0072c6 100%);	
    background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #0072c6), color-stop(1, #0072c6) );
	background:-moz-linear-gradient( center top, #0072c6 5%, #0072c6 100% );
	filter:progid:DXImageTransform.Microsoft.gradient(startColorstr="#0072c6", endColorstr="#0072c6");	
    background: -o-linear-gradient(top,#0072c6,0072c6);
	background-color:#0072c6;
	border:0px solid #7aa3c1;
	text-align:center;
	border-width:0px 0px 1px 1px;
	font-size:14px;
	font-family:Trebuchet MS;
	font-weight:bold;
	color:#ffffff;
}
.CSSTableGenerator tr:first-child:hover td{
	background:-o-linear-gradient(bottom, #0072c6 5%, #0072c6 100%);	
    background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #0072c6), color-stop(1, #0072c6) );
	background:-moz-linear-gradient( center top, #0072c6 5%, #0072c6 100% );
	filter:progid:DXImageTransform.Microsoft.gradient(startColorstr="#0072c6", endColorstr="#0072c6");	
    background: -o-linear-gradient(top,#0072c6,0072c6);
	background-color:#0072c6;
}
.CSSTableGenerator tr:first-child td:first-child{
	border-width:0px 0px 1px 0px;
}
.CSSTableGenerator tr:first-child td:last-child{
	border-width:0px 0px 1px 1px;
}
</Style>
"@

#Create Tables in split so that users can remove or add tables easily.
$start = @"
<html><div class="CSSTableGenerator">
"@
Write-Progress -Activity "ModernExchangeEnvironmentReport" -status "Writing Enviroment Initial Info"
$Table1 =@"
<table>
                    <tr><td>No.Exchange Servers</td>
                        <td>No.Databases</td>
                        <td>No.Mailboxes</td>
                        <td>No.PublicFolder Mailboxes</td>
                        <td>No.DistributionGroups</td>
                        <td>No.DynamicGroups</td>                        
                        <td>No.Contacts</td>
                        <td>No.DAG</td>
                        <td>No.Domains</td> 
                        <td>No.OrgAdmins</td> 
                        <td>Date</td>                      
                    </tr>
                    <tr><td>$($ExchangeServers.count)</td>
                        <td>$($Databases.count)</td>
                        <td>$($mailboxes.count)</td>
                        <td>$($PFmailboxes.count)</td>                                                 
                        <td>$($DistributionGroups.count)</td>
                        <td>$($DynamicGroups.count)</td>                         
                        <td>$($MailContacts.count)</td>
                        <td>$($DAG.count)</td>
                        <td>$($AcceptedDomains.count)</td>  
                        <td>$($OrgAdmins.count)</td> 
                        <td>$($Date)</td>                       
                    </tr>
</table> 
<table>
                    <tr><td>Exchange Servers</td>
                        <td>Roles</td>
                        <td>Edition</td>
                        <td>Site</td>
                        <td>Version</td>
                        <td>Operating System</td>
                        <td>SendConnectors Used</td>
                    </tr>    
"@

Write-Progress -Activity "ExchangeServers" -status "Writing ExchangeServers info"
$Table2 = 
foreach ($server in ($ExchangeServers))
             {"<tr><td>$($server.name)</td>
               <td>$($server.serverrole)</td>
               <td>$($server.edition)</td>
               <td>$($server.site.name)</td>
               <td>$($server.AdmindisplayVersion -replace "version",'')</td>
               <td>
$($windows2012above = ((Get-WmiObject -ComputerName $server.name -class Win32_OperatingSystem -ErrorAction SilentlyContinue) | Where-Object{($_.Version -like "6.*") -or ($_.version -like "10.*") -and ($_.Version -notlike "6.1.*") -and ($_.version -notlike "6.0.*")}).version.count
if ($windows2012above -eq 1)
{(Get-CimInstance -ComputerName $server.name Win32_OperatingSystem).caption -replace "Microsoft Windows Server",''})</td>
               <td>$(($SendConnectors | Where-Object{$_.SourceTransportServers -match "$($server.name)"}).identity.name)</td>               
               </tr>"}

$table2close = @"
</table>
"@

Write-Progress -Activity "Databases" -status "Writing Database info"
$Table3 =@"
<table>
                    <tr><td>Mailbox Databases</td>
                        <td>Mounted</td>
                        <td>ContentIndex</td>
                        <td>DBSize</td>                        
                        <td>Mailboxes</td>
                        <td>Master Group</td>
                        <td>Database copies</td>
                        <td>MbxRetention.Days</td>                        
                        <td>ItemRetention.Days</td>  
                        <td>CircularLogging</td>  
                        <td>Last Full Backup</td>                                          
                    </tr>
"@

$Table3data = foreach ($Database in $Databases)
             {
             $dbsize = $($Database.databasesize  -replace "\(.*",'')
             $Mountstatus = $($Database.Mounted)
              
              if($Mountstatus -eq "True")
              {
              $Mcolor = "#99FF66"
              }
              else
              {
              $Mcolor = "#FF5050"
              }
                           
              $ContentIndexStatus = $((Get-MailboxDatabaseCopyStatus "$($database.name)\$($database.servername)").contentindexstate)
              if($ContentIndexStatus -eq "Healthy")
              {
              $color = "#99FF66"
              }
              else
              {
              $color = "#FF5050"
              }
              "<tr><td>$($Database.name)</td>
               <td bgcolor = $Mcolor>$($Database.Mountedonserver)</td>
               <td bgcolor = $color>$ContentIndexStatus</td>
               <td>$dbsize</td>               
               <td>$((get-mailbox -database $Database).count)</td>
               <td>$($Database.MasterServerOrAvailabilityGroup)</td>
                <td>$($Database.databasecopies.identity.name)</td>
               <td>$($Database.MailboxRetention.days)</td>               
               <td>$($Database.DeletedItemRetention.days)</td>
               <td>$($Database.CircularLoggingEnabled -replace "False","No" -replace "True","Yes")</td>
               <td>$($database.lastfullbackup)</td>
               </tr>"
              }
             
$table3close = @"
</table>
"@

if($dags.count -ge 1)
{
Write-Progress -Activity "Dag" -status "Writing DaG info"
$Table4 =@"
<table>
                    <tr><td>DAG Name </td>
                    <td>Member Servers</td>
                        <td>DAC Mode</td>
                        <td>Witness server</td>
                        <td>Witness Directory</td>                                                                    
                    </tr>
"@


$Table4data = foreach ($dag in $dags)
              {"<tr><td>$($DAG.name)</td>
               <td>$($dag.servers.name)</td>
               <td>$($dag.DatacenterActivationMode)</td>
               <td>$($dag.Witnessserver)</td> 
               <td>$($dag.WitnessDirectory)</td>                             
               </tr>"
              }
                                         
$table4close = @"
</table>
"@

Write-Progress -Activity "Dag" -status "Writing DaG Replication info"
$Table5 =@"
<table>
                    <tr><td>Server</td>
                        <td>Check</td>
                        <td>Result</td>
                        <td>Error</td>                                                                    
                    </tr>
"@

Write-Progress -Activity "Dag" -status "Writing DaG info"

$Table5data = foreach ($dag in $dags)
              {
              foreach ($member in $((Get-DatabaseAvailabilityGroup $dag).servers.name))
              {    
              $replstatus = Test-ReplicationHealth $member
              for($i=0;$i -lt $replstatus.count;$i++) 
              {  
              if($replstatus.result[$i].value -eq "Passed")
              {
              $Rcolor = "#99FF66"
              }
              else
              {
              $Rcolor = "#FF5050"
              }     
              "<tr><td>$(($replstatus).Server[$i])</td>
               <td>$(($replstatus).Check[$i])</td>
               <td bgcolor = $Rcolor>$($replstatus.Result[$i].value)</td>
               <td>$(($replstatus).Error[$i])</td>                            
               </tr>"
              }
              }
              }
             
$table5close = @"
</table>
"@
}
else
{
Write-Progress -Activity "Dag" -status "Skipping DaG - No DAG Found"
}
$complete = @"
</div></html>
"@

#Combining All Tables.

$alltables = "$start $Table1 $Table2 $table2close $Table3 $Table3data $table3close $Table4 $Table4data $table4close $Table5 $Table5data $table5close $complete"

$Combine = ConvertTo-Html -Head $head -Body $alltables

#Saving HTML File To the local C Drive - You can modify as per your wish.
$html += $Combine
$html > $Attachment

#open the HTML File Locally
If($Openhtmllocally -eq "yes")
{
Invoke-Item "$Attachment"
}
else
{
Write-Progress -Activity "Open HTML" -status "Settings - Do not Open"
}
#Sending Email Message 

$Disclaimer = "<font color=gray><div style=font-size:8pt;font-family:Calibri,sans-serif;>
The information contained in this e-mail and any files transmitted with it are confidential and may be privileged.   Access to this e-mail by anyone other than the intended is unauthorized. If you are not the intended recipient (or responsible for delivery of the message to such person), you may not use, copy, distribute or deliver to anyone this message (or any part of its contents) or take any action in reliance on it. In such case, you should destroy this message, and notify us immediately. If you have received this email in error, please notify us immediately by e-mail or telephone and delete the e-mail from any computer. If you or your employer does not consent to internet e-mail messages of this kind, please notify us immediately. All reasonable precautions have been taken to ensure no viruses are present in this e-mail. As our company cannot accept responsibility for any loss or damage arising from the use of this e-mail or attachments we recommend that you subject these to your virus checking procedures prior to use. The views, opinions, conclusions and other information expressed in this electronic mail are not given or endorsed by the company unless otherwise indicated by an authorized representative independent of this message."

Send-mailmessage -to $EmailTo -from $EmailFrom -subject $EmailSubject -SmtpServer $SmtpServer -Body $Disclaimer -BodyAsHtml -Attachments $Attachment -ErrorAction SilentlyContinue
