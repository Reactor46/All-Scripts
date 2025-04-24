Get-Mailbox -PublicFolder | Select-Object DisplayName, PrimarySmtpAddress, Database, OrganizationalUnit

Get-PublicFolder -Recurse | Select-Object Name, ParentPath, Identity, ContentMailbox, Replicas