$ADObjects = get-aduser -filter * | ? { $_.distinguishedname -match "CN=Users,DC=ccmhmr,DC=local" }

foreach ( $ADObject in $ADObjects)
    {
    Write-HOst "Setting Inheritance on $($adobject.name)"
    $DN = $ADobject.distinguishedname
    $ou = [ADSI]"LDAP://$($DN)"
    $sec = $ou.psbase.objectSecurity

    $isProtected = $false   ## allows inheritance
    $preserveInheritance = $true  ## preserve inherited rules

    $sec.SetAccessRuleProtection($isProtected, $preserveInheritance)
    $ou.psbase.commitchanges()
    }