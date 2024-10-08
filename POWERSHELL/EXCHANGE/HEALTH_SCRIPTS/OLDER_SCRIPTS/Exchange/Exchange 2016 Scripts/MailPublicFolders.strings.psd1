# Localized	09/03/2016 06:55 AM (GMT)	303:4.80.0411 	MailPublicFolders.strings.psd1
ConvertFrom-StringData @'
###PSLOC
MailPublicFoldersUnAvailableToExport=No mail enabled folders present to export
MailPublicFoldersEnumeration=Enumerating mail enabled public folders
MailPublicFoldersEnumerationComplete=Mail public folders enumeration complete...{0} folders found.
GetExistingMailPublicFolderObjects=Retrieving existing mail public folder objects
GetExistingMailPublicFolderObjectsComplete=Retrieval of existing mail public folder objects complete...{0} folders found.
UpdateExistingMailPublicFolderObject=Updating existing mail public folder: {0}
UpdateRemovingCloudOnlyMailPublicFolderObject=WARNING: Removing Office 365 mail public folder: {0}
PublicFolderCmdletExecution=Executing Get-PublicFolder cmldet on the identified mail enabled public folders
PublicFolderCmdletExecutionComplete=Completed execution of Get-PublicFolder cmdlet...{0} folders found.
ExportMailPublicFolderObjects=Exporting mail public folder objects from Active Directory
ExportMailPublicFolderObjectsComplete=Exporting mail public folder objects from Active Directory Complete.
ImportMailPublicFolderObjects=Importing mail public folder objects into Active Directory
ImportMailPublicFolderObjectsComplete=Importing mail public folder objects into Active Directory Complete: {0} objects created, {1} objects updated and {2} objects deleted.
ErrorsFoundDuringImport=Total errors found: {0}. Please, check the error summary at '{1}' for more information.
ProgressBarTitle=Importing mail public folders...
ProcessedFolders=Items processed: {0}/{1}.
DuplicatePrimarySmtpAddressFound=The input file contains {0} mail public folder objects with PrimarySmtpAddress='{1}'.
FailedToImportDuplicateObjectsFound=Failed to import mail enabled folders because objects with duplicate PrimarySmtpAddress were found in the input file. That indicates a misconfiguration on your Exchange deployment. After fixing the problem, you can rerun the script Export-MailPublicFoldersForMigration.ps1 on your Exchange server, followed by the script Import-MailPublicFoldersForMigration.ps1 on Exchange Online to update mail enabled folders in Active Directory.
FailedToCreateMailPublicFolderEmptyPrimarySmtpAddress=Mail public folder '{0}' could not be imported because its PrimarySmtpAddress is empty.
FailedToCreateMailPublicFolder=Mail public folder could not be imported: {0}
FailedToUpdateMailPublicFolder=Mail public folder could not be updated: {0}
FailedToDeleteMailPublicFolder=Mail public folder could not be deleted: {0}
FailedToSetDisplayName=Mail public folder was imported but its DisplayName could not be set: {0}
ConfirmationTitle=Delete all mail public folders
ConfirmationQuestion=The input XML provided is empty. If you continue, all existing mail public folders in Active Directory will be deleted. Do you really want to proceed?
ConfirmationYesOption=&Yes
ConfirmationNoOption=&No
ConfirmationYesOptionHelp=Proceed and delete all existing mail public folders in Active Directory.
ConfirmationNoOptionHelp=Stop. No mail public folders will be deleted.
TimestampCsvHeader=Timestamp
IdentityCsvHeader=Identity
ErrorMessageCsvHeader=Error
###PSLOC
'@
