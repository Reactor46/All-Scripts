﻿#md c:\Transcripts
 

## Kill all inherited permissions
$acl = Get-Acl c:\Transcripts
$acl.SetAccessRuleProtection($true, $false)
 

## Grant Administrators full control
$administrators = [System.Security.Principal.NTAccount] "Administrators"
$permission = $administrators,"FullControl","ObjectInherit,ContainerInherit","None","Allow"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.AddAccessRule($accessRule)
 

## Grant everyone else Write and ReadAttributes. This prevents users from listing
## transcripts from other machines on the domain.
$everyone = [System.Security.Principal.NTAccount] "Everyone"
$permission = $everyone,"Write,ReadAttributes","ObjectInherit,ContainerInherit","None","Allow"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.AddAccessRule($accessRule)
 

## Deny "Creator Owner" everything. This prevents users from
## viewing the content of previously written files.
$creatorOwner = [System.Security.Principal.NTAccount] "Creator Owner"
$permission = $creatorOwner,"FullControl","ObjectInherit,ContainerInherit","InheritOnly","Deny"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.AddAccessRule($accessRule)
 

## Set the ACL
$acl | Set-Acl c:\Transcripts\
 

## Create the SMB Share, granting Everyone the right to read and write files. Specific
## actions will actually be enforced by the ACL on the file folder.
#New-SmbShare -Name Transcripts -Path c:\Transcripts -ChangeAccess Everyone 