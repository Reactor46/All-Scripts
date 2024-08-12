Add-MailboxPermission <username of target> -User <username to give perms> -AccessRights FullAccess


Add-ADPermission -Identity "<DisplayName of target>" -User <username to give perms> -ExtendedRights 'Send-as'