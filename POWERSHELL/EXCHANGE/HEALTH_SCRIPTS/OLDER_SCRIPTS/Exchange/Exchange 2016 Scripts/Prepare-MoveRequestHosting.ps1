﻿param([parameter(Position=0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, mandatory=$true)][string]$Identity, 
      [parameter(Position=1, mandatory=$true)][string]$RemoteForestDomainController, 
      [parameter(Position=2, mandatory=$true)][Management.Automation.PSCredential]$RemoteForestCredential, 
      [string]$LocalForestDomainController,
      [Management.Automation.PSCredential]$LocalForestCredential,
      [string]$TargetMailUserOU,
      [string]$TargetMailUserCU,
      [string]$MailboxDeliveryDomain,
      [switch]$LinkedMailUser,
      [switch]$DisableEmailAddressPolicy,
      [switch]$UseLocalObject,
      [string]$TargetExternalSmtpAddress)

begin
{
    # ---------------------------------------------------------------------------------------------------
    function findADObject($searchRoot, $filter)
    # ---------------------------------------------------------------------------------------------------
    {
        $searcher = new-object System.DirectoryServices.DirectorySearcher($searchRoot)
        $searcher.filter = $filter
        $user = $searcher.findall()

        if ($user -eq $null -or $user.count -eq 0)
        {
            return $null
        }
        elseif ($user.count -gt 1)
        {
            foreach ($usr in $user)
            {
                Write-Warning ("Object Found:" + $usr.GetDirectoryEntry().distinguishedName)
            }
            Write-Host "Tips: For Source object, please check the duplication of distinguishedName, mailNickName, displayName, objectGUID and proxyAddresses. You could parse the objectGUID in the Identity parameter to make it unique instead."
            Write-Host "Tips: For Local object, please check whether the proxyAddresses have duplicated item in these objects. You need to correct the dirty data before you could continue."
            throw "Multiple objects found in AD."
        }
        else
        {
            return $user[0].GetDirectoryEntry()
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function checkUserExist ($OU, $filter)
    # ---------------------------------------------------------------------------------------------------
    {
        $searcher = new-object System.DirectoryServices.DirectorySearcher($OU)
        $searcher.filter = $filter
        $user = $searcher.findone() 
        if ($user -eq $null -or $user.count -eq 0)
        {
            return $false
        }
        else
        {
            return $true
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function getAttributeFriendlyValue ($attribute, $attValue)
    # ---------------------------------------------------------------------------------------------------
    {
        if (($attribute -eq "msExchMailboxGuid") -or ($attribute -eq "msExchArchiveGuid"))
        {
            return ([System.Guid]($attValue)).ToString()
        }
        else
        {
            return $attValue
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function copyIfExist ($target, [array]$attriblist, $propertybag)
    # ---------------------------------------------------------------------------------------------------
    {
        foreach($att in $attriblist)
        {
            if ($propertybag.Contains($att))
            {
                $friendlyValue = getAttributeFriendlyValue $att $propertybag.Item($att).Value
                Write-Verbose "Setting $att to $friendlyValue"
                [void]($target.Put($att, $propertybag.Item($att).Value))
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function getEscapedldapFilterStr ([string]$original)
    # ---------------------------------------------------------------------------------------------------
    {
        $escape = $original.replace("\", "\5c")
        $escape = $escape.replace("(", "\28").replace(")", "\29")
        $escape = $escape.replace("&", "\26").replace("|", "\7c")
        $escape = $escape.replace("=", "\3d").replace(">", "\3e")
        $escape = $escape.replace("<", "\3c").replace("~", "\7e")
        $escape = $escape.replace("*", "\2a").replace("/", "\2f")
        $escape = $escape.replace(".", "\2e")
        return $escape
    }

    # ---------------------------------------------------------------------------------------------------
    function getEscapedDNStr ([string]$original)
    # ---------------------------------------------------------------------------------------------------
    {
        $escape = $original.replace(",", "\,")
        $escape = $escape.replace("+", "\+")
        $escape = $escape.replace("""", "\""")
        $escape = $escape.replace("#", "\#")
        $escape = $escape.replace(";", "\;")
        return $escape
    }

    # ---------------------------------------------------------------------------------------------------
    function sidToLDAPQuery([byte[]]$sid)
    # ---------------------------------------------------------------------------------------------------
    {
        foreach ($by in $sid)
        {
            $ret += "\" + $by.tostring("X")
        }
        return $ret
    }
    
    # ---------------------------------------------------------------------------------------------------    
    function MasterAccountSidIsSelf ( $srcMbxAttributes )
    # ---------------------------------------------------------------------------------------------------
    {
        if ($srcMbxAttributes.Contains("msExchMasterAccountSid"))
        {
            $master = new-object System.Security.Principal.SecurityIdentifier($srcMbxAttributes.Item("msExchMasterAccountSid").value, 0)
            if ($master.IsWellKnown("SelfSid"))
            {
                return $true
            }
        }
        return $false
    }

    # ---------------------------------------------------------------------------------------------------
    function findLocalObject ($OU, $srcuser)
    # ---------------------------------------------------------------------------------------------------
    {
        $objectClassFilter = "(&(!objectClass=computer)(|(objectClass=user)(objectClass=contact)(objectClass=group)(objectClass=msExchDynamicDistributionList)))"
        $usr = $null
        if ($srcuser.properties.Contains("msExchMasterAccountSid") -and -not (MasterAccountSidIsSelf $srcuser.properties))
        {
            $sourcesid += sidToLDAPQuery $srcuser.properties.Item("msExchMasterAccountSid").Value
            $filter = "(| (ObjectSid=$sourcesid) (msExchMasterAccountSid=$sourcesid) )"
            $filter = "(&($objectClassFilter)($filter))"

            $usr = findADObject $OU $filter
        }
        if ($usr -eq $null)
        {
            $address = $srcuser.Properties.Item("proxyAddresses")
            foreach ($addr in $address)
            {
                if ($addr.startswith("x500:", "OrdinalIgnoreCase") -or $addr.startswith("smtp:", "OrdinalIgnoreCase"))
                {
                    $addr1 = getEscapedldapFilterStr ($addr.Substring(0,4).toUpper() + $addr.Substring(4))
                    $addr2 = getEscapedldapFilterStr ($addr.Substring(0,4).toLower() + $addr.Substring(4))
                    $filterstring += "(proxyAddresses=$addr1) (proxyAddresses=$addr2)"
                }
            }
            
            $filter = "(| $filterstring)"
            $filter = "(&($objectClassFilter)($filter))"

            $usr = findADObject $OU $filter
        }
        if ($usr -eq $null)
        {
            $name = $srcuser.Properties.Item("name")
            $filter = "(&($objectClassFilter)(name=$name))"

            $usr = findADObject $OU $filter
        }
        return $usr
    }

    # ---------------------------------------------------------------------------------------------------
    function generateUniqueSAM ($ou, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $uniquesam = $srcMbxAttributes.Item('samaccountname').Value
        $retrycount = 30
        if ($uniquesam.Length -lt 20)
        {
            while ($retrycount -gt 0 -and (checkUserExist $ou "(samAccountName=$(getEscapedldapFilterStr $uniquesam))"))
            {
                $uniquesam = $srcMbxAttributes.Item("samaccountname").Value + (random)
                if ($uniquesam.length -gt 20)
                {
                    $uniquesam = $uniquesam.substring(0,20)
                }
                $retrycount = $retrycount - 1
            }
        }
        return $uniquesam
    }

    # ---------------------------------------------------------------------------------------------------
    function generateUniqueUPN ($ou, $srcMbxAttributes, $fallbacks)
    # ---------------------------------------------------------------------------------------------------
    {
        if ($srcMbxAttributes.Contains('userPrincipalName'))
        {
            $uniqueupn = $srcMbxAttributes.Item('userPrincipalName').Value
            if ($uniqueupn -match "^(.*)(@.*)$")
            {
                $postfix = $matches[2]
                $prefix  = $matches[1]
            }
            $preferedupn = ,$uniqueupn + $fallbacks
            foreach ($upn in $preferedupn)
            {
                if ($upn -ne $null)
                {
                    if ($upn.contains("@"))
                    {
                        $testupn = $upn
                    }
                    else
                    {
                        $testupn = "$upn$postfix"
                    }
                    if ($(checkUserExist $ou "(userPrincipalName=$(getEscapedldapFilterStr $testupn))") -eq $false)
                    {
                        return $testupn
                    }
                }
            }
            #try to use prefered upn, if all unsuitable, generate a new one
            while ($(checkUserExist $ou "(userPrincipalName=$(getEscapedldapFilterStr $uniqueupn))"))
            {
                $uniqueupn = "$prefix$(random)$postfix"
            }
        }
        return $uniqueupn
    }
    
    # ---------------------------------------------------------------------------------------------------
    function copyMandatoryAttributes ($newuser, $srcAttributes, $localDC)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes="displayName",
                        "Mail",
                        "mailNickName",
                        "msExchMailboxGuid",
                        "msExchArchiveGuid",
                        "msExchUserCulture",
                        "msExchArchivename",
                        "proxyAddresses"
                        
        $specialAttributes = @{  "msExchRecipientDisplayType"=0x80000006;
                                 "msExchRecipientTypeDetails"=0x80;
                                 "msExchVersion"="44220983382016";
                                 "userAccountControl"=0x202 #ACCOUNTDISABLE | NORMAL_ACCOUNT
                               }
                               
        if ($localDC -ne $null)
        {
            $specialAttributes["samaccountname"] = generateUniqueSAM $localDC $srcAttributes
            $specialAttributes["userPrincipalName"] = generateUniqueUPN $localDC $srcAttributes $newuser.cn,$specialAttributes["samaccountname"]
        }
                               
        [void](copyIfExist $newuser $copyAttributes $srcAttributes)
        
        foreach($att in $specialAttributes.getenumerator())
        {
            if ($att.value -ne $null)
            {
                Write-Verbose "Setting $($att.key) to $($att.value)"
                [void]($newuser.put($att.key, $att.value.tostring()))
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function getTargetAddress ($srcProxyAddresses, $writeErrorIfNotExist)
    # ---------------------------------------------------------------------------------------------------
    {
        foreach ($addr in $srcProxyAddresses)
        {
            if ($addr -match "^(SMTP|smtp):.*@(.*)$")
            {
                #if don't specify authoritative domains, use primary smtp address
                if (([string]::IsNullOrEmpty($MailboxDeliveryDomain) -and $addr.startswith("SMTP")) -or
                    ($matches[2] -eq $MailboxDeliveryDomain))
                {
                    Write-Verbose "Setting targetAddress to $addr"
                    return $addr
                }
            }
        }

        if ($writeErrorIfNotExist)
        {
            $errFixTip = $null
            if ([string]::IsNullOrEmpty($MailboxDeliveryDomain))
            {
                $errFixTip = "PrimaryEmailAddress exists"
            }
            else
            {
                $errFixTip = "some EmailAddress matches MailboxDeliveryDomain $MailboxDeliveryDomain"
            }

            # Terminate if fail to determine targetAddress (necessary for MEU)
            throw "Unable to determine the targetAddress for the newly created MEU. Please ensure that $errFixTip in source object."
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function disableEmailAddressPolicy ($user)
    # ---------------------------------------------------------------------------------------------------
    {
        if ($DisableEmailAddressPolicy)
        {
            Write-Verbose "Disable EmailAddressPolicy."
            try
            {
                $dcParameter = @{};
                if ($LocalForestDomainController -ne $null)
                {
                    $dcParameter = @{DomainController=$LocalForestDomainController}
                }
                Set-MailUser $user.distinguishedName.Value @dcParameter -EmailAddressPolicyEnabled:$false

            }
            catch
            {
                 # Terminate if fail to disable the EAP.
                 throw "Fail to disable the EmailAddressPolicy for MEU($($user.distinguishedName)). Error: $($Error[0])"
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function generateLegacyExchangeDN ($user)
    # ---------------------------------------------------------------------------------------------------
    {
        # Update the LegacyExchangeDN if it doesn't exist.
        if ($user.properties.Contains("LegacyExchangeDN") -eq $false)
        {
            Write-Verbose "Invoke Update-Recipient to Update LegacyExchangeDN."
            try
            {
                Update-Recipient $user.distinguishedName.Value @DomainControllerParameterSet
                #rebind ad object to retrieve new properties set by Update-Recipient
                $user.RefreshCache([array]"legacyExchangeDN")
            }
            catch
            {
                Write-Error "Error updating recipient MEU($($user.DistinguishedName)) to generate the legacyDN. Please fix the error, run the Update-Recipient task to generate the LegacyDN and this script again. Error: $($Error[0])"
                return
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function mergeLegacyDNToProxyAddress ($user, $X500ProxyAddresses, $updateImmediately, $sourceOrLocal)
    # ---------------------------------------------------------------------------------------------------
    {
        # Merge legacyDN to source object
        if ($X500ProxyAddresses.Length -gt 0)
        {
            $userFriendlyName = "Object($($user.distinguishedName))"
            if ([string]::IsNullOrEmpty($user.distinguishedName))
            {
                $userFriendlyName = "New Object" # DN is not available.
            }
            $updateRequired = $false
            foreach ($X500ProxyAddress in $X500ProxyAddresses)
            {
                if ($user.Properties.Item("proxyAddresses").tostring().toupper().Contains($X500ProxyAddress.ToUpper()) -eq $false)
                {
                    Write-Host "Appending $X500ProxyAddress to proxyAddresses of $userFriendlyName in $sourceOrLocal forest." -ForegroundColor Green
                    [void]($user.putex(3, "proxyAddresses", [array]$X500ProxyAddress))
                    $updateRequired = $true
                }
            }
            if ($updateRequired -and $updateImmediately)
            {
                try
                {
                    $user.setinfo() # Might get Access Denied
                }
                catch
                {
                    Write-Error "Error appending $X500ProxyAddress to proxyAddresses of $userFriendlyName in $sourceOrLocal forest. Please update the proxyAddresses manually, or fix the error and run this script again. Error: $($Error[0])"
                    return
                }
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function generateNewCN ($oldcn, $oudn)
    # ---------------------------------------------------------------------------------------------------
    {
            $newcn = getEscapedldapFilterStr $oldcn
            $newcn = getEscapedDNStr $newcn
            $tryCnt = 0;
            $oldcnLength = $oldcn.Length;
            while ([DirectoryServices.DirectoryEntry]::exists("LDAP://cn=$newcn,$oudn"))
            {
                if ($tryCnt -gt 10)
                {
                    throw "Unable to generate new CN, the script has tried 10 times with random string appended. You could re-run again if you're sure it could be generated.";
                }
                $tryCnt = $tryCnt + 1;

                if ($oldcnLength -ge 64)
                {
                    # if old CN legth is already maximum (64), we can't generate new one.
                    throw "Unable to generate new CN. The original CN $newcn exists and its length >= 64, thus we can't append random string at the end to generate new one.";
                }

                $ranStr = (random).ToString()
                if ($oldcnLength + $ranStr.Length -gt 64)
                {
                    $ranStr = $ranStr.subString(0, 64 - $oldcnLength);
                }
                $newcn = getEscapedldapFilterStr ($oldcn + $ranStr)
                $newcn = getEscapedDNStr $newcn
            }
            return $newcn;
    }

    # ---------------------------------------------------------------------------------------------------
    function createMailUserAccount ($localDC, $ou, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        try{
            $newcn = generateNewCN $srcMbxAttributes.Item("cn").value $ou.distinguishedname

            [void]($newuser = $ou.create("user", "cn=$newcn"))
            
            copyMandatoryAttributes $newuser $srcMbxAttributes $localDC

            #additional operations for proxyaddresses and targetaddress

            if ($srcMbxAttributes.Contains("LegacyExchangeDN"))
            {
                $X500proxyAddr = "x500:" + $srcMbxAttributes.Item("LegacyExchangeDN").value
                mergeLegacyDNToProxyAddress $newuser ([array]$X500proxyAddr) $false "Local"
            }

            $srcProxyAddresses = $srcMbxAttributes.Item("proxyAddresses")
            $targetAddress = getTargetAddress $srcProxyAddresses $true
            if ($targetAddress -ne $null)
            {
                [void]($newuser.put("targetAddress", $targetAddress))
            }

            [void]($newuser.SetInfo())

            $newuser.RefreshCache([array]"distinguishedName")

            return $newuser
        }
        catch
        {
            # Terminate the script if fail to create the new user.
            throw "Error creating mailuser CN=$newcn,$($ou.distinguishedname) in local forest or setting its mandatory attributes. Error: $($Error[0])"
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function copyGalySyncAttributes ($user, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes= "C",
                         "Co",
                         "countryCode",
                         "Company",
                         "Department",
                         "facsimileTelephoneNumber",
                         "givenName",
                         "homePhone",
                         "Info",
                         "Initials",
                         "L",
                         "Mobile",
                         "msExchAssistantName",
                         "msExchHideFromAddressLists",
                         "otherHomePhone",
                         "otherTelephone",
                         "Pager",
                         "physicalDeliveryOfficeName",
                         "postalCode",
                         "Sn",
                         "St",
                         "streetAddress",
                         "telephoneAssistant",
                         "telephoneNumber",
                         "Title"

        copyIfExist $user $copyAttributes $srcMbxAttributes
    }

    # ---------------------------------------------------------------------------------------------------
    function copyE2k7OptionalAttributes ($user, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes= #"Cn",
                         "Comment",
                         "deletedItemFlags",
                         "delivContLength",
                         "departmentNumber",
                         "Description",
                         "Division",
                         "employeeID",
                         "employeeNumber",
                         "employeeType",
                         "homePostalAddress",
                         "internationalISDNNumber",
                         "ipPhone",
                         "Language",
                         "localeID",
                         "mAPIRecipient",
                         "middleName",
                         "msDS-PhoneticCompanyName",
                         "msDS-PhoneticDepartment",
                         "msDS-PhoneticDisplayName",
                         "msDS-PhoneticFirstName",
                         "msDS-PhoneticLastName",
                         "msExchBlockedSendersHash",
                         "msExchELCExpirySuspensionEnd",
                         "msExchELCExpirySuspensionStart",
                         "msExchELCMailboxFlags",
                         "msExchExternalOOFOptions",
                         "msExchMessageHygieneFlags",
                         "msExchMessageHygieneSCLDeleteThreshold",
                         "msExchMessageHygieneSCLJunkThreshold",
                         "msExchMessageHygieneSCLQuarantineThreshold",
                         "msExchMessageHygieneSCLRejectThreshold",
                         "msExchMDBRulesQuota",
                         "msExchPoliciesExcluded",
                         "msExchSafeRecipientsHash",
                         "msExchSafeSendersHash",
                         "msExchUMSpokenName",
                         "O",
                         "otherFacsimileTelephoneNumber",
                         "otherIpPhone",
                         "otherMobile",
                         "otherPager",
                         "preferredDeliveryMethod",
                         "personalPager",
                         "personalTitle",
                         "Photo",
                         "pOPCharacterSet",
                         "pOPContentFormat",
                         "postalAddress",
                         "postOfficeBox",
                         "primaryInternationalISDNNumber",
                         "primaryTelexNumber",
                         "showInAdvancedViewOnly",
                         "Street",
                         "terminalServer",
                         "textEncodedORAddress",
                         "thumbnailLogo",
                         "thumbnailPhoto",
                         "url",
                         "userCert",
                         "userCertificate",
                         "userSMIMECertificate",
                         "wWWHomePage"
        foreach ($i in 1..15)
        {
            $copyAttributes += "extensionAttribute$i";
        }


        copyIfExist $user $copyAttributes $srcMbxAttributes
    }

    # ---------------------------------------------------------------------------------------------------
    function findCorrespondingADObject ($targetOU, $DN, $srcDomain)
    # ---------------------------------------------------------------------------------------------------
    {
        $cn = "$DN".substring(0, "$DN".indexof(",DC="))
        $srcreferenceobject = $srcDomain.children.find($cn)
        $usr = $null
        if ($srcreferenceobject -ne $null)
        {
            if ($srcreferenceobject.Properties.Contains("legacyExchangeDN"))
            {
                $legexch = getEscapedldapFilterStr $srcreferenceobject.Properties.Item("legacyExchangeDN")
                $addrfilter = "(proxyAddresses=x500:$legexch) (proxyAddresses=X500:$legexch)"
            }
            $address = $srcreferenceobject.Properties.Item("proxyAddresses")
            foreach ($addr in $address)
            {
                if ($addr.startswith("x500:", "OrdinalIgnoreCase"))
                {
                    $addrfilter += "(legacyExchangeDN=$(getEscapedldapFilterStr $addr.substring(5)))"
                }
                if ($addr.startswith("smtp:", "OrdinalIgnoreCase") -or $addr.startswith("x500:", "OrdinalIgnoreCase"))
                {
                    $addr1 = getEscapedldapFilterStr ($addr.Substring(0,4).toUpper() + $addr.Substring(4))
                    $addr2 = getEscapedldapFilterStr ($addr.Substring(0,4).toLower() + $addr.Substring(4))
                    $addrfilter += "(proxyAddresses=$addr1) (proxyAddresses=$addr2)"
                }
            }
            if ([string]::IsNullOrEmpty($addrfilter) -eq $false)
            {
                $filter = "(| $addrfilter)"
                $objectClassFilter = "(&(!objectClass=computer)(|(objectClass=user)(objectClass=contact)(objectClass=group)(objectClass=msExchDynamicDistributionList)))"

                $filterWithObjectClass = "(&($objectClassFilter)($filter))"
                $usr = findADObject $targetOU $filterWithObjectClass

                if ($usr -eq $null)
                {
                    #user not found, try find with loose condition. Because link/backlink may be in other object type.
                    $usr = findADObject $targetOU $filter
                }
            }

            return $usr
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function setLinkedAttribute ($attribname, $backlinkname, $targetOU, $user, $srcMbxAttributes, $srcDomain)
    # ---------------------------------------------------------------------------------------------------
    {
        if ($srcMbxAttributes.contains($attribname))
        {
            foreach ($dn in $srcMbxAttributes.item($attribname))
            {
                try
                {
                    $corobj = findCorrespondingADObject $targetOU $dn $srcDomain
                    if ($corobj -eq $null)
                    {
                        Write-Warning "Cannot find corresponding object for $dn in current forest. `'$attribname`' not set."
                    }
                    else
                    {
                        Write-Verbose "Setting $attribname to $($corobj.properties.item('distinguishedname'))"
                        if (($attribname -eq "Manager") -or ($attribname -eq "managedBy") -or ($attribname -eq "altRecipient"))
                        {
                            $user.put($attribname, $($corobj.properties.item('distinguishedname')))
                        }
                        else
                        {
                            $user.putex(3, $attribname, [array]"$($corobj.properties.item('distinguishedname'))")
                        }
                    }
                }
                catch
                {
                    Write-Warning "Error updating $($user.distinguishedName)   Attribute: $attribname! Attribute Not Set! Error: $($Error[0])"
                }
            }
        }
        
        #find backlink from source MBX, set it on corresponding user in target
        if ($srcMbxAttributes.contains($backlinkname))
        {
            foreach ($dn in $srcMbxAttributes.item($backlinkname))
            {
                try
                {
                    $corobj = findCorrespondingADObject $targetOU $dn $srcDomain
                    if ($corobj -eq $null)
                    {
                        Write-Warning "Cannot find corresponding object for $dn in current forest. `'$attribname`' not updated."
                    }
                    else
                    {
                        if (($attribname -eq "Manager") -or ($attribname -eq "managedBy") -or ($attribname -eq "altRecipient"))
                        {
                            $corobj.Put($attribname, $($user.properties.item("distinguishedname")))
                        }
                        else
                        {
                            $corobj.PutEx(3, $attribname, [array]"$($user.properties.item("distinguishedname"))")
                        }
                        $corobj.SetInfo()
                        Write-Host "Updating $($corobj.distinguishedName)   Attribute: $attribname" -ForegroundColor Green
                    }
                }
                catch
                {
                    Write-Warning "Error updating $($corobj.distinguishedName)   Attribute: $attribname! Attribute Not Set! Error: $($Error[0])"
                }
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function setLinkedAttributes ($targetOU, $user, $srcMbxAttributes, $srcDomain)
    # ---------------------------------------------------------------------------------------------------
    {
        setLinkedAttribute "altRecipient" "altRecipientBL" $targetOU $user $srcMbxAttributes $srcDomain
        if ($user.properties.contains("altRecipient") -and $srcMbxAttributes.contains("deliverAndRedirect"))
        {
            $user.put("deliverAndRedirect", "$($srcMbxAttributes.item('deliverAndRedirect'))".toupper())
        }
        
        setLinkedAttribute "Manager" "directReports" $targetOU $user $srcMbxAttributes $srcDomain
        
        setLinkedAttribute "publicDelegates" "publicDelegatesBL"  $targetOU $user $srcMbxAttributes $srcDomain
        
        setLinkedAttribute "member" "memberOf"  $targetOU $user $srcMbxAttributes $srcDomain

        setLinkedAttribute "managedBy" "managedObjects"  $targetOU $user $srcMbxAttributes $srcDomain

        setLinkedAttribute "msExchCoManagedByLink" "msExchCoManagedObjectsBL"  $targetOU $user $srcMbxAttributes $srcDomain
    }

    # ---------------------------------------------------------------------------------------------------
    function copyLinkedMailboxTypeAttributes ($user, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes = @()
        $valuedAttributes = @{ }
        
        $accountenable = ($srcMbxAttributes.Item("UserAccountControl").tostring() -band 0x2) -eq 0
        if (-not $accountenable -and (MasterAccountSidIsSelf $srcMbxAttributes))
        {
            $valuedAttributes["msExchRecipientDisplayType"] = $user.properties.Item("msExchRecipientDisplayType").value -bor 2
        }
        else
        {
            $valuedAttributes["msExchRecipientDisplayType"] = 0xC0000006
            if ($srcMbxAttributes.Contains("msExchMasterAccountSid"))
            {
                $copyAttributes += "msExchMasterAccountSid"
            }
            elseif ($srcMbxAttributes.Contains("objectSid"))
            {
                $valuedAttributes["msExchMasterAccountSid"] = $srcMbxAttributes.Item("objectSid").Value
            }
            #this can also be done by carefully arrange "msExchMasterAccountSid" and "objectSid"
            #in the list, avoid the trouble of nested branching. but it's not worth the maintainence effort
        }
        
        [void](copyIfExist $user $copyAttributes $srcMbxAttributes)
    
        foreach($att in $valuedAttributes.getenumerator())
        {
            Write-Verbose "Setting $($att.key) to $($att.value)"
            [void]($user.put($att.key, $att.value))
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function copySpecialMailboxTypeAttributes ($user, $srcMbxAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
    #Mailbox Type Constants
        $ROOMMAILBOX      = 16
        $EQUIPMENTMAILBOX = 32
    #End Mailbox Type Constants
        $copyAttributes = "msExchResourceCapacity",
                          "msExchResourceDisplay",
                          "msExchResourceMetaData",
                          "msExchResourceSearchProperties"
                          
        $valuedAttributes = @{ }
        
        if ($srcMbxAttributes.Contains("msExchRecipientTypeDetails"))
        {
            $typedetail = $user.ConvertLargeIntegerToInt64($srcMbxAttributes.Item("msExchRecipientTypeDetails").Value)
            if (($typedetail -band $ROOMMAILBOX) -ne 0)
            {
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000706
            }
            elseif (($typedetail -band $EQUIPMENTMAILBOX) -ne 0)
            {
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000806
            }
            else
            {
                return
            }

            [void](copyIfExist $user $copyAttributes $srcMbxAttributes)
            
            foreach($att in $valuedAttributes.getenumerator())
            {
                Write-Verbose "Setting $($att.key) to $($att.value)"
                [void]($user.put($att.key, $att.value))
            }
        }
    }

    # ---------------------------------------------------------------------------------------------------
    function createMEUAndCopyAttrs ($localDC, $localOU, $srcDC, $srcObject)
    # ---------------------------------------------------------------------------------------------------
    {
        $srcAttributes = $srcObject.properties
        $newuser = createMailUserAccount $localDC $localOU $srcAttributes

        #mandatory attributes are all set. go with optional attributes
        copyGalySyncAttributes $newuser $srcAttributes
        copyE2k7OptionalAttributes $newuser $srcAttributes
        setLinkedAttributes $localdc $newuser $srcAttributes $srcdc
        copySpecialMailboxTypeAttributes $newuser $srcAttributes

        if ($LinkedMailUser)
        {
            copyLinkedMailboxTypeAttributes $newuser $srcAttributes
        }

        try
        {
            [void]($newuser.SetInfo())
        }
        catch
        {
            Write-Error "Fail to update attributes of local MEU $($newuser.DistinguishedName). $($Error[0])"
            return
        }

        [void](disableEmailAddressPolicy($newuser))
        [void](generateLegacyExchangeDN($newuser))

        #syncback Legacy Exchange DN
        if ($newuser.properties.Contains("LegacyExchangeDN"))
        {
            $X500proxyAddr = "x500:" + $newuser.properties.Item("LegacyExchangeDN")
            mergeLegacyDNToProxyAddress $srcObject ([array]$X500proxyAddr) $true "Source"
        }

        $Global:movecount++
        "Preparation for $Identity done."
    }
    
    # ---------------------------------------------------------------------------------------------------
    function forceMergeObject ($recipienttype, $localOU, $localCU, $localusr, $srcObject)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes = "msExchMailboxGUID",
                          "msExchArchiveGUID",
                          "msExchArchiveName",
                          "userPrincipalName"

        $X500proxyAddrsToUpdateSrcObj = @()
        $localUserOriginalLegacyDNToX500 = $null
        if ($localusr.Properties.Contains("LegacyExchangeDN"))
        {
            # Original user's legacyDN will be merged to source object.
            $localUserOriginalLegacyDNToX500 = "x500:" + $localusr.Properties.Item("LegacyExchangeDN")
            $X500proxyAddrsToUpdateSrcObj += $localUserOriginalLegacyDNToX500
        }

        if ($recipienttype -eq 'MailUser')
        {
            [void](disableEmailAddressPolicy($localusr))

            Write-Verbose "Merging Mailbox properties to local MailUser"
            [void](copyIfExist $localusr $copyAttributes $srcObject.properties)
            [void]($localusr.put("msExchVersion", "44220983382016"))
            $logindisabled = ($srcObject.userAccountControl.Value -band 0x2) -ne 0 #AccountDisabled
            if ($LinkedMailUser -and $logindisabled)
            {
                copyLinkedMailboxTypeAttributes $localusr $srcObject.properties
            }

            # Merge source object's legacyDN to local user's proxyAddresses.
            if ($srcObject.properties.Contains("LegacyExchangeDN"))
            {
                $X500proxyAddr = "x500:" + $srcObject.properties.Item("LegacyExchangeDN")
                mergeLegacyDNToProxyAddress $localusr ([array]$X500proxyAddr) $false "Local"
            }
            
            if (-not $localusr.Properties.Contains("msExchOURoot"))
            {
                [void]$localusr.put("msExchOURoot", $localOU.distinguishedname)
            }
            
            if (-not $localusr.Properties.Contains("msExchCU"))
            {
                [void]$localusr.put("msExchCU", $localCU.distinguishedname)
            }

            try
            {
                [void]($localusr.SetInfo())  # Might get Access Denied.
            }
            catch
            {
                Write-Error "Error updating local MEU($($localusr.DistinguishedName)) attributes from source object. Please fix the error and run this script again. Error: $($Error[0])"
            }
        }
        elseif ($recipienttype -eq 'MailContact')
        {
            Write-Verbose "Creating MailUser with same attributes as local MailContact"
            
            $srcMbxAttributes = $srcObject.Properties
            $ContactAttributes = $localusr.Properties

            $newcn = generateNewCN $srcMbxAttributes.Item("cn").value $localou.distinguishedname
            $newuser = $localOU.create("user", "cn=$newcn")
            
            copyMandatoryAttributes $newuser $ContactAttributes
            copyGalySyncAttributes $newuser $ContactAttributes
            copyE2k7OptionalAttributes $newuser $ContactAttributes
            copySpecialMailboxTypeAttributes $newuser $ContactAttributes

            if ($LinkedMailUser)
            {
                copyLinkedMailboxTypeAttributes $newuser $ContactAttributes
            }
            
            [void](copyIfExist $newuser "targetAddress" $ContactAttributes)

            $srcProxyAddresses = $srcMbxAttributes.Item("proxyAddresses")
            $writeErrorIfTargetAddrNotExist = (-not $ContactAttributes.Contains("targetAddress"))
            $srcMbxTargetAddress = getTargetAddress $srcProxyAddresses $writeErrorIfTargetAddrNotExist
            if ($srcMbxTargetAddress -ne $null)
            {
                [void]$newuser.put("targetAddress", $srcMbxTargetAddress)
            }

            # Create the User first, so that we could update the link/backlink of the user before deleting local user.
            try
            {
                [void]($newuser.setinfo())
            }
            catch
            {
                # Terminiate the script if fail to create the new MEU.
                throw "Error creating MailUser cn=$newcn,$($localOU.distinguishedName) with with attributes from source MBX($($srcObject.distinguishedName)). Error: $($Error[0])"
            }
            $newuser.RefreshCache([array]"distinguishedName")
            Write-Host -ForegroundColor green "New MEU $($newuser.distinguishedname) created successfully."

            # backlink needs to be set after user is creaed correctly.
            setLinkedAttributes $localdc $newuser $ContactAttributes $localdc

            Write-Host -ForegroundColor red "Deleteing $($localusr.distinguishedname)"
            try
            {
                deleteLocalUser($localusr)
            }
            catch
            {
                # Terminate the script if fail to delete the object.
                throw "Error deleting local MailContact($($localusr.distinguishedname)). Error: $($Error[0])"
            }

            Write-Verbose "Updating MailUser with with attributes from source MBX($($srcObject.distinguishedName))."

            if ($LinkedMailUser)
            {
                copyLinkedMailboxTypeAttributes $newuser $srcMbxAttributes
            }

            $copyAttributes += "sAMAccountName"
            [void](copyIfExist $newuser $copyAttributes $srcMbxAttributes)
            [void]($newuser.put("msExchVersion", "44220983382016"))

            # Merge source object and original local user's LegacyDN to NEW local user's proxyAddresses.
            $X500proxyAddrsToUpdateLocalObj = @()
            if ($localUserOriginalLegacyDNToX500 -ne $null)
            {
                $X500proxyAddrsToUpdateLocalObj += $localUserOriginalLegacyDNToX500
            }
            if ($srcObject.properties.Contains("LegacyExchangeDN"))
            {
                $srcObjectLegacyDNX500 = "x500:" + $srcObject.properties.Item("LegacyExchangeDN")
                if (($localUserOriginalLegacyDNToX500 -ne $null) -and ($localUserOriginalLegacyDNToX500.ToString().ToUpper() -ne $srcObjectLegacyDNX500.ToString().ToUpper()))
                {
                    $X500proxyAddrsToUpdateLocalObj += $srcObjectLegacyDNX500
                }
            }
            mergeLegacyDNToProxyAddress $newuser $X500proxyAddrsToUpdateLocalObj $false "Local"

            [System.Threading.Thread]::Sleep(500)

            try
            {
                [void]($newuser.setinfo())
            }
            catch
            {
                Write-Error "Error updating MailUser cn=$newcn,$($localOU.distinguishedName). Error: $($Error[0])"
            }

            [void](disableEmailAddressPolicy($newuser))
            [void](generateLegacyExchangeDN($newuser))

            # new user's legacyDN will be merged to source object's proxyAddresses
            if ($newuser.properties.Contains("LegacyExchangeDN"))
            {
                $newUserLegacyDNX500 = "x500:" + $newuser.properties.Item("LegacyExchangeDN")
                if (($localUserOriginalLegacyDNToX500 -ne $null) -and ($localUserOriginalLegacyDNToX500.ToString().ToUpper() -ne $newUserLegacyDNX500.ToString().ToUpper()))
                {
                    $X500proxyAddrsToUpdateSrcObj += $newUserLegacyDNX500
                }
            }
        }

        mergeLegacyDNToProxyAddress $srcObject $X500proxyAddrsToUpdateSrcObj $true "Source"

        "Preparation for $Identity done. Local recipient info Merged."
        $Global:movecount++
    }
    
    # ---------------------------------------------------------------------------------------------------
    # Delete specified local user through S.DS.P instead of ADSI. This is in order to fix bug #270965
    function deleteLocalUser ($localusr)
    # ---------------------------------------------------------------------------------------------------
    {
        $server = getMailContactOriginatingServer $localusr
        $cnx = createLocalLdapConnection $server
        $deleteRequest = [System.DirectoryServices.Protocols.DeleteRequest]($localusr.distinguishedName.ToString())
        [void]($cnx.SendRequest($deleteRequest))
        $cnx.Dispose()
    }

    # ---------------------------------------------------------------------------------------------------
    function getMailContactOriginatingServer($localusr)
    # ---------------------------------------------------------------------------------------------------
    {
        $mailContact = Get-MailContact $localusr.distinguishedName.Value @DomainControllerParameterSet
        $mailContactServer = $mailContact.OriginatingServer
        $server = $LocalForestDomainController
        if (($mailContactServer -ne $null) -and (-not [string]::IsNullOrEmpty($mailContactServer.ToString())))
        {
            $server = $mailContactServer.ToString()
        }
        Write-Verbose "Local MailContact's OriginatingServer is $server"
        return $server
    }

    # ---------------------------------------------------------------------------------------------------
    # Create a LdapConnection to the local server.
    function createLocalLdapConnection($server)
    # ---------------------------------------------------------------------------------------------------
    {
        $directoryId = [System.DirectoryServices.Protocols.LdapDirectoryIdentifier]($server)
        if ($LocalForestCredential -ne $null)
        {
            $cnx = new-object "System.DirectoryServices.Protocols.LdapConnection"($directoryId, $LocalForestCredential.GetNetworkCredential())
        }
        else
        {
            $cnx = new-object "System.DirectoryServices.Protocols.LdapConnection"($directoryId)
        }
            
        $cnx.SessionOptions.AutoReconnect = $true
        $cnx.SessionOptions.ReferralChasing = [System.DirectoryServices.Protocols.ReferralChasingOptions]::None
        $cnx.SessionOptions.Signing = $true
        $cnx.AuthType = [System.DirectoryServices.Protocols.AuthType]::Kerberos
        $cnx.Bind()
        
        return $cnx
    }
    
#=========================================================================================================
#                                         Initialize code
#=========================================================================================================
    [void]([System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.Protocols"))
    
    $usr = $RemoteForestCredential.UserName
    $pwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($RemoteForestCredential.Password))
    
    if ($LocalForestCredential -ne $null)
    {
        $localusr = $LocalForestCredential.UserName
        $localpwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($LocalForestCredential.Password))
    }
    $Global:movecount = 0

    $srcdc   = New-Object DirectoryServices.DirectoryEntry("LDAP://$RemoteForestDomainController", $usr, $pwd)
    $DomainControllerParameterSet = @{}
    if ($srcdc.guid -eq $null)
    {
        #guid not present, consider src unavailable
        throw "Source Domain controller unavailable or authentication failed."
    }

    try {
        if ($LocalForestCredential -eq $null -and [string]::IsNullorEmpty($LocalForestDomainController))
        {
            $localdc = [ADSI]""
        }
        elseif ($LocalForestCredential -ne $null -and $LocalForestDomainController -ne $null)
        {
            $localdc = New-Object DirectoryServices.DirectoryEntry("LDAP://$LocalForestDomainController", $localusr, $localpwd)
            $DomainControllerParameterSet = @{ DomainController=$LocalForestDomainController; Credential=$LocalForestCredential }
        }
        else
        {
            throw "LocalForestCredential and LocalForestDomainController need to be specified at the same time"
        }

        $escapedtargetou = $null
        $filterObjectClass = $null
        if ([string]::IsNullOrEmpty($TargetMailUserOU))
        {
            # By default, the target OU is in Users container
            $escapedtargetou = "Users"
            $filterObjectClass = "Container"
        }
        else
        {
            $escapedtargetou = $TargetMailUserOU
            $filterObjectClass = "organizationalUnit"
        }
        $OUfilter = "(& (ObjectClass=$filterObjectClass)" +
                    "   (| (name=$escapedtargetou)" +
                    "      (distinguishedname=$escapedtargetou)))"
        $localOU =  findADObject $localdc $OUfilter
        if ($localOU -eq $null)
        {
            throw "Cannot find specified OU or Container: $TargetMailUserOU"
        }
        
        $escapedtargetcu = $null
        $filterObjectClass = $null
        $localCU = $null
        if (-not [string]::IsNullOrEmpty($TargetMailUserCU))
        {
            $cnPath = "CN=Configuration," + $localdc.distinguishedName.ToString()
            if ($LocalForestCredential -eq $null -and [string]::IsNullorEmpty($LocalForestDomainController))
            {
                $localcn = [ADSI]"LDAP://$cnPath"
            }
            elseif ($LocalForestCredential -ne $null -and $LocalForestDomainController -ne $null)
            {
                $localcn = New-Object DirectoryServices.DirectoryEntry("LDAP://$LocalForestDomainController/$cnPath", $localusr, $localpwd)
                $DomainControllerParameterSet = @{ DomainController=$LocalForestDomainController; Credential=$LocalForestCredential }
            }
            
            $filterObjectClass = "container"
       
            $CUfilter = "(& (ObjectClass=$filterObjectClass)" +
                        "   (| (name=$TargetMailUserCU)" +
                        "      (distinguishedname=$TargetMailUserCU)))"
            $localCU =  findADObject $localcn $CUfilter
            if ($localCU -eq $null)
            {
                throw "Cannot find specified OU or Container: $TargetMailUserCU"
            }
        }
    }
    catch
    {
        throw "Error looking up local OU, Error Msg: $($Error[0])"
    }
}

process
{    
    $escapedIdentity = getEscapedldapFilterStr $Identity
    $filterDN =   "(& (objectClass=user)(!objectClass=computer)" +
                  "   (distinguishedName=$escapedIdentity))"

    $filterParm = "(& (objectClass=user)(!objectClass=computer)" +
                  "   ( (| (mailnickname=$escapedIdentity)" + 
                  "        (cn=$escapedIdentity)" +
                  "        (proxyAddresses=SMTP:$escapedIdentity)" +
                  "        (proxyAddresses=smtp:$escapedIdentity)" +
                  "        (proxyAddresses=X500:$escapedIdentity)" +
                  "        (proxyAddresses=x500:$escapedIdentity)" +
                  "        (objectGUID=$escapedIdentity)" +
                  "        (displayname=$escapedIdentity))))"

    try
    {
        $srcObject = findADObject $srcdc $filterParm

        if ($srcObject -eq $null)
        {
            $srcObject = findADObject $srcdc $filterDN
        
            if ($srcObject -eq $null)
            {
                Write-Error "Error looking up source MBX $identity in source forest."
                return
            }
        }
    }
    catch
    {
        Write-Error "Faile to lookup the source object. Error $($Error[0])"
        return
    }
    
    if (-not $srcObject.properties.contains("mailNickName") -or -not $srcObject.properties.contains("msExchHomeServerName"))
    {
        Write-Error "Source Object $($srcObject.distinguishedName) found, but it is not a Mailbox!."
        return
    }
    
    $accountenable = ($srcObject.properties.Item("UserAccountControl").tostring() -band 0x2) -eq 0
    
    if (-not $accountenable -and -not $srcObject.properties.contains("msExchMasterAccountSid"))
    {
        Write-Error "Source Mailbox is invalid because it is disabled but did not set msExchMasterAccountSid."
        return
    }
    
    try
    {
        $userPrincipalName = $srcObject.Properties["UserPrincipalName"]
    
        if (-not [string]::IsNullOrEmpty($TargetMailUserOU))
        {
            $ou = [ADSI]"LDAP://$TargetMailUserOU"
            $localusr = findLocalObject $ou $srcObject
        }
        else 
        {
            $localusr = findLocalObject $localdc $srcObject
        }
        if ((-not $localusr.Properties.Contains("msExchCU")) -and (-not $localusr.Properties.Contains("msExchOURoot")))
        {
            $localusr.Properties["msExchCU"].Add($TargetMailUserCU)
            $localusr.Properties["msExchOURoot"].Add($TargetMailUserOU)
            $localusr.CommitChanges()
        }

        if ((Get-MailUser $localusr.distinguishedName.ToString()) -eq $null)
        {
            if (-not [string]::IsNullOrEmpty($TargetExternalSmtpAddress))
            {
                Enable-MailUser $localusr.distinguishedName.ToString() -ExternalEmailAddress:$TargetExternalSmtpAddress
				Set-MailUser $localusr.distinguishedName.ToString() -UserPrincipalName:$userPrincipalName.ToString()
            }
            else
            {
                Enable-MailUser $localusr.distinguishedName.ToString() -ExternalEmailAddress:$userPrincipalName.ToString()
            }
        }
    }
    catch
    {
        Write-Error "Error processing $identity, Mailbox not ready to move! Error message: $($error[0])"
        return
    }
    if ($localusr -eq $null)
    {
        try
        {
            #local recipient not exist, source object found, proceed the MEU creation process
            createMEUAndCopyAttrs $localdc $localOU $srcDC $srcObject
        }
        catch
        {
            Write-Error "Error while creating MEU. Error:$($Error[0])";
            return
        }
    }
    else
    {
        Write-Verbose "Local ad account with dupplicate proxy addresses found: $($localusr.distinguishedName)"
        try
        {
            $recipienttype = (get-recipient $localusr.distinguishedname.value @DomainControllerParameterSet).RecipientType
            if ($recipienttype -eq 'MailUniversalDistributionGroup' -or $recipienttype -eq 'UserMailbox' -or $recipienttype -eq 'DynamicDistributionGroup')
            {
                    write-error "Cannot create mail enabled user because an existing mailbox user or mail enabled group already has the same proxy addresses/MasterAccountSid."
            }
            elseif ($recipienttype -eq 'MailUser' -or $recipienttype -eq 'MailContact')
            {
                if ($UseLocalObject)
                {
                    forceMergeObject $recipienttype $localOU $localCU $localusr $srcObject
                }
                else
                {
                    write-error ("Cannot create mail enabled user because an existing mail enabled user " +
                                 "or contact already has the same proxy addresses/MasterAccountSid. Please rerun the script with " + 
                                 "'-UseLocalObject' if you want to convert the existing email enabled user or contact to " +
                                 "a mail enabled user that is ready for online mailbox move.")
                }
            }
            else
            {
                write-error "Cannot create mail enabled user because an existing object with type $recipienttype already has the same proxy addresses/MasterAccountSid."
            }
        }
        catch
        {
            Write-Error "Fail to prepare local existed (same Proxyaddress or Masteraccoutsid) user $($localusr.distinguishedName) for move request. Error: $($Error[0])"
            return
        }
    }
}

end
{
    Write-Host -ForegroundColor Black -BackgroundColor Green "$movecount mailbox(s) ready to move."
}
# SIG # Begin signature block
# MIIdvAYJKoZIhvcNAQcCoIIdrTCCHakCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUhc6lUzfOHHy4cSOzPHFvac/n
# AYCgghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI4
# WhcNMTcwNjMwMTkyMTI4WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAipCth86FRu1y
# rtsPu2NLSV7nv6A/oVAcvGrG7VQwRn+wGlXrBf4nyiybHGn9MxuB9u4EMvH8s75d
# kt73WT7lGIT1yCIh9VC9ds1iWfmxHZtYutUOM92+a22ukQW00T8U2yowZ6Gav4Q7
# +9M1UrPniZXDwM3Wqm0wkklmwfgEEm+yyCbMkNRFSCG9PIzZqm6CuBvdji9nMvfu
# TlqxaWbaFgVRaglhz+/eLJT1e45AsGni9XkjKL6VJrabxRAYzEMw4qSWshoHsEh2
# PD1iuKjLvYspWv4EBCQPPIOpGYOxpMWRq0t/gqC+oJnXgHw6D5fZ2Ccqmu4/u3cN
# /aAt+9uw4wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFHbWEvi6BVbwsceywvljICto
# twQRMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBABbNYMMt3JjfMAntjQhHrOz4aUk970f/hJw1jLfYspFpq+Gk
# W3jMkUu3Gev/PjRlr/rDseFIMXEq2tEf/yp72el6cglFB1/cbfDcdimLQD6WPZQy
# AfrpEccCLaouf7mz9DGQ0b9C+ha93XZonTwPqWmp5dc+YiTpeAKc1vao0+ru/fuZ
# ROex8Zd99r6eoZx0tUIxaA5sTWMW6Y+05vZN3Ok8/+hwqMlwgNR/NnVAOg2isk9w
# ox9S1oyY9aRza1jI46fbmC88z944ECfLr9gja3UKRMkB3P246ltsiH1fz0kFAq/l
# 2eurmfoEnhg8n3OHY5a/Zzo0+W9s1ylfUecoZ4UwggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBMIwggS+AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB1jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUy/Y6HainXCh6ds0fCbkRDbEfv5YwdgYKKwYB
# BAGCNwIBDDFoMGagPoA8AFAAcgBlAHAAYQByAGUALQBNAG8AdgBlAFIAZQBxAHUA
# ZQBzAHQASABvAHMAdABpAG4AZwAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAdUN6es/hXvY81c7h
# uGPg89YCyVjU1zqMeabG/mlWCT3DOxVfkoK5j23nXmDGtFG9Q1BSB8NNDkowKw7/
# NOzI8smnBWCAe7CnxS+36RlDl/G0AQhlhWxmtq9ERfCaRKe2u9hPxAr5thZszVBW
# eD9VBUmGG3nSn22EfWd7AaX0GMF1Q3nl8BbCewwYGhDKm1pVpmrvljb9JXGHoXlL
# rn58FmyK4YNe2JOdGFXBTtrnIOcRadZkoze/wUfJkee/h0AaX7jOHRzz2WgBJPHe
# chOEsAnPq6nPeqhT3KUj4FHgXwsC4gybvqUPZC4VatrEdLdZrl1lC9kE0Ab0O4Hu
# 7Zklj6GCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0ECEzMAAACZqsWBn4yifYoAAAAAAJkwCQYFKw4DAhoFAKBd
# MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkw
# MzE4NDQwMVowIwYJKoZIhvcNAQkEMRYEFL7Baxct/BNHw0u6wrgvH3jHYz4OMA0G
# CSqGSIb3DQEBBQUABIIBAGEhh8TK6WRPmaYb0PMBplwQ2HIr4hfmKad9MtDsXWu2
# CR8Wi+Uq8JDNYzm9EpKKrnnUzjtWs/44XCDg7Gdpf1XG+Awiuz+/ZO+dZeV9Tj60
# yJOQlZpchS7xDkSYXq1DjXy6ZjoWXpOAlt8pyN9hEjNDmQXVmg1Dc0qShifSGvCB
# xbU2xxXYMKTgbcBvwxSaavaCMFbr5+H+XIiUAT94DdZMxAVBz1g9HCbJqNcRxN8u
# 5Z4qwKBeCkIR6hmTlMEm5lFQJRsk50uRTRLmHpue8agP1Pwml1Cx+u2X6lSxLw85
# nN1XTKqy4D8ul5CXkQcuT307GNqx0a5zZuTyRNY41NM=
# SIG # End signature block
