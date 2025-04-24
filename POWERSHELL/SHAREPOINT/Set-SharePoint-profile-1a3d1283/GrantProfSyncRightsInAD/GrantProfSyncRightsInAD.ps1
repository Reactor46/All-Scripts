param($profSyncAccountName) 

<#
written by Ingo Karstein (http://blog.karstein-consulting.com)
v1.0 - USE IT AT YOUR OWN RISK!

Download: http://gallery.technet.microsoft.com/Set-SharePoint-profile-1a3d1283
#>

#debug
#$profSyncAccountName = "spprofsync@kc-dev.com"

if( [string]::IsNullOrEmpty($profSyncAccountName) ) {
    Write-Error "No account name specified!"
    return
}

$profsyncAccount = new-object System.Security.Principal.NTAccount($profSyncAccountName);
$profSyncAccountSid =  $profsyncAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value


[System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices") | out-null

$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$domainEntry = $domain.GetDirectoryEntry()


$domainPath = $domainEntry.Path
$domainConfigurationPath = "LDAP://$($domain.name)/CN=Configuration,$($domainEntry.distinguishedName)"

$domainPath, $domainConfigurationPath | % {
    $entry = $_

    $container = new-object System.DirectoryServices.DirectoryEntry($entry)
    $securityDescriptor = $container.ObjectSecurity

    $foundProfSyncInACL = $null
    $foundProfSyncInACL = ($securityDescriptor.Access | ? { 
        #write-host ($_.IdentityReference)
        $a = new-object System.Security.Principal.NTAccount($_.IdentityReference)
        $aSid = $a.Translate([System.Security.Principal.SecurityIdentifier]).Value
        $aSid.Equals($profSyncAccountSid)
    }) 

    $foundProfSyncInACL = $foundProfSyncInACL -ne $null 

    if( $foundProfSyncInACL ) {
        Write-Error "Account already in ACL"
        return 
    }


    #Look "http://msdn.microsoft.com/en-us/library/cc223512.aspx" for GUID
    #  1131f6aa-9c07-11d1-f79f-00c04fc2dcd2 => "DS-Replication-Get-Changes"

    $extRightAccessRule = new-object System.DirectoryServices.ExtendedRightAccessRule($profsyncAccount, "Allow", [guid]"1131f6aa-9c07-11d1-f79f-00c04fc2dcd2");
    $propertyReadAccessRule = new-object System.DirectoryServices.ActiveDirectoryAccessRule($profsyncAccount, "ReadProperty", "Allow")
    $genericExecAccessRule = new-object System.DirectoryServices.ActiveDirectoryAccessRule($profsyncAccount, "GenericExecute", "Allow")

    $securityDescriptor.AddAccessRule($extRightAccessRule)
    $securityDescriptor.AddAccessRule($propertyReadAccessRule)
    $securityDescriptor.AddAccessRule($genericExecAccessRule)

    $container.CommitChanges()
}