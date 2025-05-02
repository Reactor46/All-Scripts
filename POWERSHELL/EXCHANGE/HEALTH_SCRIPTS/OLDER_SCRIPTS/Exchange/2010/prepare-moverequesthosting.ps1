param([parameter(Position=0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, mandatory=$true)][string]$Identity, 
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
# MIIacAYJKoZIhvcNAQcCoIIaYTCCGl0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUhc6lUzfOHHy4cSOzPHFvac/n
# AYCgghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA5MDQyMTQy
# MDlaFw0xMzAzMDQyMTQyMDlaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6pElsEPsi
# nGWiFpg7y2Fi+nQprY0GGdJxWBmKXlcNaWJuNqBO/SJ54B3HGmGO+vyjESUWyMBY
# LDGKiK4yHojbfz50V/eFpDZTykHvabhpnm1W627ksiZNc9FkcbQf1mGEiAAh72hY
# g1tJj7Tf0zXWy9kwn1P8emuahCu3IWd01PZ4tmGHmJR8Ks9n6Rm+2bpj7TxOPn0C
# 6/N/r88Pt4F+9Pvo95FIu489jMgHkxzzvXXk/GMgKZ8580FUOB5UZEC0hKo3rvMA
# jOIN+qGyDyK1p6mu1he5MPACIyAQ+mtZD+Ctn55ggZMDTA2bYhmzu5a8kVqmeIZ2
# m2zNTOwStThHAgMBAAGjggENMIIBCTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU3lHcG/IeSgU/EhzBvMOzZSyRBZgwHwYDVR0jBBgwFoAUyxHoytK0FlgB
# yTcuMxYWuUyaCh8wVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAu
# Y3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMDgtMzEtMjAxMC5jcnQw
# DQYJKoZIhvcNAQEFBQADggEBACqk9+7AwyZ6g2IaeJxbxf3sFcSneBPRF1MoCwwA
# Qj84D4ncZBmENX9Iuc/reomhzU+p4LvtRxD+F9qHiRDRTBWg8BH/2pbPZM+B/TOn
# w3iT5HzVbYdx1hxh4sxOZLdzP/l7JzT2Uj9HQ8AOgXBTwZYBoku7vyoDd3tu+9BG
# ihcoMaUF4xaKuPFKaRVdM/nff5Q8R0UdrsqLx/eIHur+kQyfTwcJ7SaSbrOUGQH4
# X4HnrtqJj39aXoRftb58RuVHr/5YK5F/h9xGH1GVzMNiobXHX+vJaVxxkamNViAs
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggS6MIIDoqADAgEC
# AgphAo5CAAAAAAAfMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQTAeFw0xMjAxMDkyMjI1NThaFw0xMzA0MDkyMjI1NThaMIGzMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYD
# VQQLEx5uQ2lwaGVyIERTRSBFU046RjUyOC0zNzc3LThBNzYxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCW7I5HTVTCXJWA104LPb+XQ8NL42BnES8BTQzY0UYvEEDeC6RQ
# UhKIC0N6LT/uSG5mx5HmA8pu7HmpaiObzWKezWqkP+ejQ/9iR6G0ukT630DBhVR+
# 6KCnLEMjm1IfMjX0/ppWn41jd3swngozhXIbykrIzCXN210RLsewjPGPQ0hHBbV6
# IAvl8+/BuvSz2M04j/shqj0KbYUX0MrnhgPAM4O1JcTMWpzEw9piJU1TJRRhj/sb
# 4Oz3R8aAReY1UyM2d8qw3ZgrOcB1NQ/dgUwhPXYwxbKwZXMpSCfYwtKwhEe7eLrV
# dAPe10sZ91PeeNqG92GIJjO0R8agVIgVKyx1AgMBAAGjggEJMIIBBTAdBgNVHQ4E
# FgQUL+hGyGjTbk+yINDeiU7xR+5IwfIwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2
# +7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBY
# BggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUE
# DDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAc/99Lp3NjYgrfH3jXhVx
# 6Whi8Ai2Q1bEXEotaNj5SBGR8xGchewS1FSgdak4oVl/de7G9TTYVKTi0Mx8l6uT
# dTCXBx0EUyw2f3/xQB4Mm4DiEgogOjHAB3Vn4Po0nOyI+1cc5VhiIJBFL11FqciO
# s3xybRAnxUvYb6KoErNtNSNn+izbJS25XbEeBedDKD6cBXZ38SXeBUcZbd5JhaHa
# SksIRiE1qHU2TLezCKrftyvZvipq/d81F8w/DMfdBs9OlCRjIAsuJK5fQ0QSelzd
# N9ukRbOROhJXfeNHxmbTz5xGVvRMB7HgDKrV9tU8ouC11PgcfgRVEGsY9JHNUaeV
# ZTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEFBQAwXzETMBEG
# CgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsG
# A1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTEw
# MDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENvZGUgU2lnbmlu
# ZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCycllcGTBkvx2a
# YCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPVcgDbNVcKicqu
# IEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlcRdyvrT3gKGiX
# GqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZC/6SdCnidi9U
# 3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgGhVxOVoIoKgUy
# t0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdcpReejcsRj1Y8
# wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTL
# EejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYBBAGCNxUBBAUC
# AwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy82C0wGQYJKwYB
# BAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBWJ5flJRP8KuEK
# U5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNyb3NvZnQuY29t
# L3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3JsMFQGCCsGAQUF
# BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcNAQEFBQADggIB
# AFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGjI8x8UJiAIV2s
# PS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbNLeNK0rxw56gN
# ogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y4k74jKHK6BOl
# kU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnpo1hW3ZsCRUQv
# X/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6H0q70eFW6NB4
# lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20OE049fClInHLR
# 82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8Z4L5UrKNMxZl
# Hg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9GuwdgR2VgQE6w
# QuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXrilUEnacOTj5XJ
# jdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEvmtzjcT3XAH5i
# R9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCCA++gAwIBAgIK
# YRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPyLGQBGRYDY29t
# MRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1MzA5WhcNMjEw
# NDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZUSNQrc7dGE4k
# D+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cOBJjwicwfyzMk
# h53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn1yjcRlOwhtDl
# KEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3U21StEWQn0gA
# SkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG7bfeI0a7xC1U
# n68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMBAAGjggGrMIIB
# pzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A+3b7syuwwzWz
# DzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1UdIwSBkDCBjYAU
# DqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
# GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0BxMuZTBQBgNV
# HR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
# CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
# Y3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwTq86+e4+4LtQS
# ooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+jwoFyI1I4vBT
# Fd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwYTp2Oawpylbih
# OZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPfwgphjvDXuBfr
# Tot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5ZlizLS/n+YWG
# zFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8csu89Ds+X57H21
# 46SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUwZuhCEl4ayJ4i
# IdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHudiG/m4LBJ1S2
# sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9La9Zj7jkIeW1
# sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g74TKIdbrHk/J
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggS0
# MIIEsAIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHWMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBTL9jodqKdcKHp2zR8JuRENsR+/ljB2BgorBgEEAYI3AgEMMWgwZqA+gDwAUABy
# AGUAcABhAHIAZQAtAE0AbwB2AGUAUgBlAHEAdQBlAHMAdABIAG8AcwB0AGkAbgBn
# AC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDAN
# BgkqhkiG9w0BAQEFAASCAQBUlNWKQolLMcDSyXvoOpPjSkGVOKcR8eo5Y/a/ofgP
# b26IO93XjC0hkx6Ew6x7UdPyGtMvTAbUhoBB665Y94yC5Ureemt55QbyJTEfIJVE
# +evIBkIDBK487+LtLYQ0eOzsMVsCjpeXQn7g5+ot57IUz2FpWAYxwgGXvGKI8vZX
# SmznBBmv3z4NCBN/ydDVr6PKNBTfnEJAbHJg+2SNHCKl7YrnEgDWJIffqFWRCVnS
# cEsn6TR0TwJeoianU38FnfdSB0W4suV44WaqVSOe8TSs8pemmAIg4SGLCb3xuII9
# SVPoEQ9xkEeDZx8NRuTijXaD8gLD4DXGrvFk0yXheOQgoYICHzCCAhsGCSqGSIb3
# DQEJBjGCAgwwggIIAgEBMIGFMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQIKYQKO
# QgAAAAAAHzAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAc
# BgkqhkiG9w0BCQUxDxcNMTMwMjA1MDYzNzI1WjAjBgkqhkiG9w0BCQQxFgQUN+gk
# 5EJUfxRixp1O0eyfPpSNCzwwDQYJKoZIhvcNAQEFBQAEggEASpCj8N9VXVKk7UnM
# 00sAuN0Mz9clFAuHHl6QWElI9d+5h1uUD7MNuLw6W8LHAqA3rcYQOACoxbopzZdt
# fqax6BeWChWHyxCnd5r/xIgMuFF3bkk8VzRilphCtvKMThy0j8nORq2oCB8MQ8se
# GAd0vB87Xdm/OXrRfRZlIJa8f1UuDCljAT9o27zhnhYvEj6wd2EjBrjJLMptUbiL
# HfaICZVw8zjEVOl1BvHsYVRU8rNY45ytK7TDRrKGg6TU1D1NLNjj0uSioau/0WdX
# IJSUwGFY0PBSuEj4qmw8eM4ms8UBGU+orlT8Erme6cNsSYXegUOVmcCBv7f2uNeu
# BUSKvw==
# SIG # End signature block
