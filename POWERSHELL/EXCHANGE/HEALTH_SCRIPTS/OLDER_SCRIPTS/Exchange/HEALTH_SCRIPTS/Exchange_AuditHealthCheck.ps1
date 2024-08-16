# ===================================================================
# NAME: Exchange_AuditHealthCheck.ps1
# AUTHOR: PRIGENT Nicolas (www.get-cmd.com)
# DATE: 11/16/2015
#
# Exchange Audit and Health check (Mailbox, Database, ...)
# 
# COMMENTS: 
# v1.0 - 11/16/2015 - N.PRIGENT : Creation
#
# ===================================================================

### To be modified
$MBXServerName = "ServerName"
$ToMail = "EmailAddr"
$FromMail = "CheckExchange@domain.com"
$SmtpServer = "mail.domain.com"
$MailEnabled = $False
$Date = Get-Date
$File = "C:\Exchange_AuditHealthCheck_" + $Date.Tostring('HHmm-MMddyyyy') + ".htm"

Add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
### Script variables
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$MBXServerName/PowerShell/ -Authentication Kerberos 
#Import-PSSession $Session
$tabGlobalMBXListing = @()
$tabEXCRole = @()
$tabDBDate = @()
$tabMBXListing = @()
$tabPubFolders = @()

### CSS style
$css= "<style>"
$css= $css+ "BODY{ text-align: center; background-color:white;}"
$css= $css+ "TABLE{    font-family: 'Lucida Sans Unicode', 'Lucida Grande', Sans-Serif;font-size: 12px;margin: 10px;width: 100%;text-align: center;border-collapse: collapse;border-top: 7px solid #004466;border-bottom: 7px solid #004466;}"
$css= $css+ "TH{font-size: 13px;font-weight: normal;padding: 1px;background: #cceeff;border-right: 1px solid #004466;border-left: 1px solid #004466;color: #004466;}"
$css= $css+ "TD{padding: 1px;background: #e5f7ff;border-right: 1px solid #004466;border-left: 1px solid #004466;color: #669;hover:black;}"
$css= $css+  "TD:hover{ background-color:#004466;}"
$css= $css+ "</style>"

### Roles
$EXCRole = Get-ExchangeServer
foreach ($EXCRol in $EXCRole) {
$objEXCRole = new-object Psobject
$objEXCRole | Add-member -Name "Name" -Membertype "Noteproperty" -Value $EXCRol.Name
$objEXCRole | Add-member -Name "Site" -Membertype "Noteproperty" -Value $EXCRol.Site
$objEXCRole | Add-member -Name "Server Role" -Membertype "Noteproperty" -Value $EXCRol.ServerRole
$objEXCRole | Add-member -Name "Edition" -Membertype "Noteproperty" -Value $EXCRol.Edition
$objEXCRole | Add-member -Name "Version" -Membertype "Noteproperty" -Value $EXCRol.AdminDisplayVersion
$tabEXCRole += $objEXCRole
}

### MBX DB
$DBDate = Get-MailboxDatabase -Status | select ServerName,Name,DatabaseSize, Mounted, LastFullBackup, Circularloggingenabled, MountAtStartup,organization
$tabGlobalDBDate = @()
foreach ($DBDat in $DBDate) {
$objDBDate = new-object Psobject
$tabDBDate = @()
$DBName = $DBDat.Name
$CountMBX = (Get-Mailbox -Database $DBName -resultsize unlimited).count
$objDBDate | Add-member -Name "Server Name" -Membertype "Noteproperty" -Value $DBDat.ServerName
$objDBDate | Add-member -Name "Name" -Membertype "Noteproperty" -Value $DBName
$objDBDate | Add-member -Name "Database Size" -Membertype "Noteproperty" -Value $DBDat.DatabaseSize
$objDBDate | Add-member -Name "Count Mailbox" -Membertype "Noteproperty" -Value $CountMBX
$objDBDate | Add-member -Name "Mounted" -Membertype "Noteproperty" -Value $DBDat.Mounted
$objDBDate | Add-member -Name "Last Full Backup" -Membertype "Noteproperty" -Value $DBDat.LastFullBackup
$objDBDate | Add-member -Name "Mount At Startup ?" -Membertype "Noteproperty" -Value $DBDat.MountAtStartup
$objDBDate | Add-member -Name "Circular Logging Enabled ?" -Membertype "Noteproperty" -Value $DBDat.Circularloggingenabled
$objDBDate | Add-member -Name "Organization" -Membertype "Noteproperty" -Value $DBDat.organization
$tabDBDate += $objDBDate
$tabGlobalDBDate += $tabDBDate
}

### List MBX
$MBXListing = Get-MailboxStatistics -Server $MBXServerName | Sort-Object TotalItemSize -Descending | select-object -First 25 |select DisplayName,DeletedItemCount,TotalDeletedItemSize,ItemCount,TotalItemSize,LastLogonTime
$tabGlobalMBXListing = @()
foreach ($MBXList in $MBXListing) {
$objMBXListing = new-object Psobject
$tabMBXListing = @()
$objMBXListing | Add-member -Name "Display Name" -Membertype "Noteproperty" -Value $MBXList.Displayname
$objMBXListing | Add-member -Name "Deleted Item Count" -Membertype "Noteproperty" -Value $MBXList.DeletedItemCount
$objMBXListing | Add-member -Name "Total Deleted Item Size" -Membertype "Noteproperty" -Value $MBXList.TotalDeletedItemSize
$objMBXListing | Add-member -Name "Total Item Count" -Membertype "Noteproperty" -Value $MBXList.ItemCount
$objMBXListing | Add-member -Name "Total Item Size" -Membertype "Noteproperty" -Value $MBXList.TotalItemSize
$objMBXListing | Add-member -Name "Last Logon Time" -Membertype "Noteproperty" -Value $MBXList.LastLogonTime
$tabMBXListing += $objMBXListing
$tabGlobalMBXListing += $tabMBXListing
}

### Public Folders
$PubFolders = Get-PublicFolderStatistics | select-object -First 25
$tabPubFolders = @()
foreach ($PubFolder in $PubFolders) {
$objPubFolders = new-object Psobject
$objPubFolders | Add-member -Name "Name" -Membertype "Noteproperty" -Value $PubFolder.Name
$objPubFolders | Add-member -Name "Total Item Count" -Membertype "Noteproperty" -Value $PubFolder.ItemCount
$objPubFolders | Add-member -Name "Total Item Size" -Membertype "Noteproperty" -Value $PubFolder.TotalItemSize
$objPubFolders | Add-member -Name "Last Modification Time" -Membertype "Noteproperty" -Value $PubFolder.LastModificationTime
$tabPubFolders += $objPubFolders
}

### Count DG and members per DG
$DGroups = @(Get-DistributionGroup -ResultSize Unlimited)
$tabGlobalDG = @()
foreach ($DG in $DGroups) { 
$count = @(Get-ADGroupMember -Recursive $dg.DistinguishedName).Count 
$objDG = New-Object PSObject
$tabDG = @()
$objDG | Add-Member NoteProperty -Name "Group Name" -Value $DG.Name 
$objDG | Add-Member NoteProperty -Name "DN" -Value $DG.distinguishedName 
$objDG | Add-Member NoteProperty -Name "SMTP Addr" -Value $DG.PrimarySmtpAddress
$objDG | Add-Member NoteProperty -Name "Hidden" -Value $DG.HiddenFromAddressListsEnabled
$objDG | Add-Member NoteProperty -Name "Member Count" -Value $count 
$objDG | Add-Member NoteProperty -Name "When changed" -Value $DG.WhenChanged
$tabDG += $objDG 
$tabGlobalDG += $tabDG
}

### Test Mail Flow
$MailFlow = Test-Mailflow -TargetMailboxServer $MBXServerName
foreach ($MailFl in $MailFlow) { 
$objMailFl = New-Object PSObject
$tabMailFl = @()
$objMailFl | Add-Member NoteProperty -Name "MBX Server Name" -Value $MBXServerName
$objMailFl | Add-Member NoteProperty -Name "Result" -Value $MailFl.TestMailflowResult
$objMailFl | Add-Member NoteProperty -Name "Message Latency Time" -Value $MailFl.MessageLatencyTime
$tabMailFl += $objMailFl
}

### Audit logs in the last 24H
$search = Search-AdminAuditLog -StartDate ((Get-Date).AddHours(-24)) -EndDate (Get-Date)
$tabGlobalAudit = @()
foreach ($sea in $search) { 
$objAudit = New-Object PSObject
$tabAudit = @()
$objAudit | Add-Member NoteProperty -Name "Caller" -Value $sea.Caller
$objAudit | Add-Member NoteProperty -Name "Cmdlet Name" -Value $sea.CmdletName
$objAudit | Add-Member NoteProperty -Name "From which server ?" -Value $sea.OriginatingServer
$objAudit | Add-Member NoteProperty -Name "Run Date" -Value $sea.RunDate
$objAudit | Add-Member NoteProperty -Name "Object Modified" -Value $sea.ObjectModified
$tabAudit += $objAudit
$tabGlobalAudit += $tabAudit
}

# sort tables
$tabGlobalMBXListing = $tabGlobalMBXListing | Sort-Object "Total Item Size" -descending
$tabPubFolders = $tabPubFolders | Sort-Object "Total Item Size" -descending
$tabGlobalDG = $tabGlobalDG | Sort-Object "Member Count" -descending

# Add to the body
$body = "<center><h1>Exchange Audit and Health Check</h1></center>" 
$body += "<center>By Nicolas PRIGENT <a href='http://www.get-cmd.com'>[www.get-cmd.com]</a></center>"
$body += "<h4>Servers Informations</h4>" 
$body += $tabEXCRole | ConvertTo-Html -Head $css 

$body += "<h4>Databases Informations</h4>" 
$body += $tabGlobalDBDate | ConvertTo-Html -Head $css 

$body += "<h4>List Mailbox By Size (Top 25 largest mailboxes) </h4>" 
$body += $tabGlobalMBXListing | ConvertTo-Html -Head $css

$body += "<h4>List Distribution Group with total members</h4>" 
$body += $tabGlobalDG | ConvertTo-Html -Head $css

$body += "<h4>List Public Folders (Top 25 largest PFs)</h4>" 
$body += $tabPubFolders | ConvertTo-Html -Head $css

$body += "<h4>Test Mail Flow</h4>" 
$body += $tabMailFl | ConvertTo-Html -Head $css

$body += "<h4>Audit Logs in the last 24H</h4>" 
$body += $tabGlobalAudit | ConvertTo-Html -Head $css

# If enabled send mail
if ($MailEnabled) {
    send-mailmessage -to $ToMail -from $FromMail -subject "Exchange audit and health check" -body ($body | out-string) -BodyAsHTML -SmtpServer $SmtpServer
} else {
    $body | Out-File $File
}

#Remove-PSSession $Session