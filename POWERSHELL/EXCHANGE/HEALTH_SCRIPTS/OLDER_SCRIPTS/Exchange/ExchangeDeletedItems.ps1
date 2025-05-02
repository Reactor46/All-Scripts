#Get the list of mailboxes (Display Name, Mailbox Size, DB and MailboxGUID)
	Write-Host "Gathering a list of Mailboxes (including Display Name, Total Item Size, Database and MailboxGUID).."
	$mailboxes = Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | Sort-Object TotalItemSize -Descending | Select-Object DisplayName,TotalItemSize,Database,MailboxGuid
 
#Output file
	$yyyyMMdd = Get-Date -format "yyyyMMdd"
	$txtOutput = "C:\ExchangeReport_$($yyyyMMdd).txt"
	
	#Deletes the output file if it exists
		If (Test-Path $txtOutput){
			Remove-Item $txtOutput
		}
	#Creates the first line of the output file
		# Splitting them with a "|" instead of a comma. It'll be easier to work on it later on
		# due the commas in the FolderSize result.
		Add-Content -Path $txtOutput -Value "DisplayName|TotalItemSize|Database|DeletedItemsSize|UserAccountControl|OrganizationalUnit|MailboxGUID"
	
#Count the number of mailboxes based on $mailboxes
	$numberOfmailboxes = $mailboxes.count
 
#Initialise the counter
	$counter = 1
 
$mailboxes | ForEach-Object {
	
	#Get MailboxGUID
		$identity = $_.MailboxGUID.guid
	#Get Archive DB
		$archiveDB = (Get-Mailbox -identity $identity | Select-Object ArchiveDatabase).Name
	#Get OU
		$OU = (Get-Mailbox -identity $identity | Select-Object OrganizationalUnit).OrganizationalUnit
	#Get DB
		$DB = $_.database.name
	#Get DisplayName
		$DisplayName = $_.displayname
	#Get TotalItemSize
		$TotalItemSize = $_.TotalItemSize.Value
	#Get UserAccountControl to report if the account is disabled.
		$UserAccountControl = (Get-User -identity $identity | Select-Object UserAccountControl).UserAccountControl
	#Get FolderSize for DeletedItems Folder scope. If the variable is empty or it's an array, try to get it based (again) on its Scope
	#but this time limiting the result to a few known language (there might be a few more). If also this time the variable is an array
	#or it's empty, then try to get the size searching through all folders (again, based on known languages).
	#This should speed up the process a lot if the folder is found.
		$FolderSize = $FolderSize = (get-mailboxfolderstatistics -identity $identity -FolderScope DeletedItems | Select-Object FolderSize).FolderSize
			If ($FolderSize -is [system.array] -OR $FolderSize -eq $NULL)
			{
				$FolderSize = (get-mailboxfolderstatistics -identity $identity -FolderScope DeletedItems | where {$_.Name -eq "Deleted Items" -OR $_.Name -eq "Itens Excluídos" -OR $_.Name -eq "Éléments supprimés" -OR $_.Name -eq "Itens Eliminados" -OR $_.Name -eq "Elementos eliminados" -OR $_.Name -eq "Gelöschte Objekte" -OR $_.Name -eq "Elementy usuniete" -OR $_.Name -eq "Törölt elemek" -OR $_.Name -eq "Gelöschte Elemente" -OR $_.Name -eq "Mensagens excluídas"} | Select-Object FolderSize).FolderSize
				
				If ($FolderSize -is [system.array] -OR $FolderSize -eq $NULL)
				{
					$FolderSize = (get-mailboxfolderstatistics -identity $identity | where {$_.Name -eq "Deleted Items" -OR $_.Name -eq "Itens Excluídos" -OR $_.Name -eq "Correos eliminados" -OR $_.Name -eq "Éléments supprimés" -OR $_.Name -eq "Itens Eliminados" -OR $_.Name -eq "Elementos eliminados" -OR $_.Name -eq "Gelöschte Objekte" -OR $_.Name -eq "Elementy usuniete" -OR $_.Name -eq "Törölt elemek" -OR $_.Name -eq "Gelöschte Elemente" -OR $_.Name -eq "Mensagens excluídas"} | Select-Object FolderSize).FolderSize
				}
			}
			
	
		$output = "$($DisplayName)|$($TotalItemSize)|$($DB)|$($FolderSize)|$($UserAccountControl)|$($OU)|$($identity)"
		Add-Content -Path $txtOutput -Value "$($output)"
		Write-Host "$($counter) of $($numberOfmailboxes)"
		Write-Host "$($output)" #Displays the output on screen (I like to know where the script is at and what it is doing)
		Write-Host ""
	
	$counter = $counter+1
	
}