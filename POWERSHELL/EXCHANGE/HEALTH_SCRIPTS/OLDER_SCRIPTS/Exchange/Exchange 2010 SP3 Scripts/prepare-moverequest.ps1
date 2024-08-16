param([parameter(Position=0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, mandatory=$true)][string]$Identity, 
      [parameter(Position=1, mandatory=$true)][string]$RemoteForestDomainController, 
      [parameter(Position=2, mandatory=$true)][Management.Automation.PSCredential]$RemoteForestCredential, 
      [string]$LocalForestDomainController,
      [Management.Automation.PSCredential]$LocalForestCredential,
      [string]$TargetMailUserOU,
      [string]$MailboxDeliveryDomain,
      [switch]$LinkedMailUser,
      [switch]$DisableEmailAddressPolicy,
      [switch]$UseLocalObject,
      [switch]$OverwriteLocalObject)

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
    function copyBasicAttributes ($newuser, $srcAttributes)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes="displayName",
                        "Mail",
                        "mailNickName",
                        "msExchMailboxGuid",
                        "msExchArchiveGuid",
                        "msExchUserCulture",
                        "msExchArchivename"

        [void](copyIfExist $newuser $copyAttributes $srcAttributes)
    }


    # ---------------------------------------------------------------------------------------------------
    function copyMandatoryAttributes ($newuser, $srcAttributes, $localDC)
    # ---------------------------------------------------------------------------------------------------
    {
        copyBasicAttributes $newuser $srcAttributes

        # Handle proxyAddresses specially (only copied when the local object is created at first).
        [void](copyIfExist $newuser "proxyAddresses" $srcAttributes)

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
    function generateNewCN ($localDC, $oldcn, $oudn)
    # ---------------------------------------------------------------------------------------------------
    {
            $newcn = getEscapedldapFilterStr $oldcn
            $newcn = getEscapedDNStr $newcn
            $tryCnt = 0;
            $oldcnLength = $oldcn.Length;
            $localPath = $localdc.Path;
            # Usage of this script may not provide the LocalForestDomainController parameter, in which case, the $localdc.Path is empty.
            if ([string]::IsNullorEmpty($localPath))
            {
                $localPath = "LDAP:/";
            }
            while ([DirectoryServices.DirectoryEntry]::exists("$localPath/cn=$newcn,$oudn"))
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
            $newcn = generateNewCN $localDC $srcMbxAttributes.Item("cn").value $ou.distinguishedname

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
        $copyAttributes = "msExchResourceCapacity",
                          "msExchResourceDisplay",
                          "msExchResourceMetaData",
                          "msExchResourceSearchProperties"
                          
        $valuedAttributes = @{ }

        $isResource = $false;

        if ($srcMbxAttributes.Contains("msExchRecipientTypeDetails"))
        {
            $ROOMMAILBOX      = 16
            $EQUIPMENTMAILBOX = 32

            $typedetail = $user.ConvertLargeIntegerToInt64($srcMbxAttributes.Item("msExchRecipientTypeDetails").Value)
            if (($typedetail -band $ROOMMAILBOX) -ne 0)
            {
                $isResource = $true;
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000706
            }
            elseif (($typedetail -band $EQUIPMENTMAILBOX) -ne 0)
            {
                $isResource = $true;
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000806
            }
        }

        if ((-not $isResource) -and ($srcMbxAttributes.Contains("msExchRecipientDisplayType")))
        {
            $ROOMMAILBOX      = 0x80000706;
            $EQUIPMENTMAILBOX = 0x80000806;

            $recipientDisplayType = $srcMbxAttributes.Item("msExchRecipientDisplayType").Value;
            if ($recipientDisplayType -eq $ROOMMAILBOX)
            {
                $isResource = $true;
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000706;
            }
            elseif ($recipientDisplayType -eq $EQUIPMENTMAILBOX)
            {
                $isResource = $true;
                $valuedAttributes["msExchRecipientDisplayType"] = 0x80000806;
            }
        }

        if ($isResource)
        {
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
    function forceMergeObject ($recipienttype, $localOU, $localusr, $localDC, $srcObject, $srcDC)
    # ---------------------------------------------------------------------------------------------------
    {
        $copyAttributes = "msExchMailboxGUID",
                          "msExchArchiveGUID",
                          "msExchArchiveName"

        $X500proxyAddrsToUpdateSrcObj = @()
        $localUserOriginalLegacyDNToX500 = $null
        if ($localusr.Properties.Contains("LegacyExchangeDN"))
        {
            # Original user's legacyDN will be merged to source object.
            $localUserOriginalLegacyDNToX500 = "x500:" + $localusr.Properties.Item("LegacyExchangeDN")
            $X500proxyAddrsToUpdateSrcObj += $localUserOriginalLegacyDNToX500
        }

        $userToUpdate = $null;

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

            # Set the proxyAddresses first before merging legacyDN.
            if ($OverwriteLocalObject)
            {
                [void](copyIfExist $localusr "proxyAddresses" $srcObject.properties)
            }

            # Merge source object's legacyDN to local user's proxyAddresses.
            if ($srcObject.properties.Contains("LegacyExchangeDN"))
            {
                $X500proxyAddr = "x500:" + $srcObject.properties.Item("LegacyExchangeDN")
                mergeLegacyDNToProxyAddress $localusr ([array]$X500proxyAddr) $false "Local"
            }

            try
            {
                [void]($localusr.SetInfo())  # Might get Access Denied.
            }
            catch
            {
                Write-Error "Error updating local MEU($($localusr.DistinguishedName)) attributes from source object. Please fix the error and run this script again. Error: $($Error[0])"
            }

            $userToUpdate = $localusr;
        }
        elseif ($recipienttype -eq 'MailContact')
        {
            Write-Verbose "Creating MailUser with same attributes as local MailContact"
            
            $srcMbxAttributes = $srcObject.Properties
            $ContactAttributes = $localusr.Properties

            $newcn = generateNewCN $localDC $srcMbxAttributes.Item("cn").value $localou.distinguishedname
            $newuser = $localOU.create("user", "cn=$newcn")
            
            copyMandatoryAttributes $newuser $ContactAttributes
            copyGalySyncAttributes $newuser $ContactAttributes
            copyE2k7OptionalAttributes $newuser $ContactAttributes

            if ($LinkedMailUser)
            {
                # Copy necessary values (like msExchMasterAccountSid) from contact first, it will be overridden
                # by following one from source mbx.
                copyLinkedMailboxTypeAttributes $newuser $ContactAttributes
            }

            # This must come after copyLinkedMailboxTypeAttributes since it will initialize the value.
            copySpecialMailboxTypeAttributes $newuser $ContactAttributes

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

            # Set the proxyAddresses first before merging legacyDN.
            if ($OverwriteLocalObject)
            {
                [void](copyIfExist $newuser "proxyAddresses" $srcMbxAttributes)
            }

            $copyAttributes += "sAMAccountName",
                               "userPrincipalName"
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

            $userToUpdate = $newuser;
        }

        mergeLegacyDNToProxyAddress $srcObject $X500proxyAddrsToUpdateSrcObj $true "Source"

        # Overwrite local object from source object if -OverwriteLocalObject is specified.
        if ($OverwriteLocalObject)
        {
            Write-Verbose "OverwriteLocalObject specified. Updating MailUser with with attributes from source MBX($($srcObject.distinguishedName))."

            $srcAttributes = $srcObject.properties
            copyBasicAttributes $userToUpdate $srcAttributes
            copyGalySyncAttributes $userToUpdate $srcAttributes
            copyE2k7OptionalAttributes $userToUpdate $srcAttributes
            setLinkedAttributes $localdc $userToUpdate $srcAttributes $srcdc
            copySpecialMailboxTypeAttributes $userToUpdate $srcAttributes

            try
            {
                [void]($userToUpdate.setinfo())
            }
            catch
            {
                Write-Error "Error updating MailUser($($userToUpdate.DistinguishedName)). Error: $($Error[0])"
            }
        }

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

    if ($OverwriteLocalObject -and (-not $UseLocalObject))
    {
        throw "Please specify -UseLocalObject when you specify -OverwriteLocalObject parameter."
    }

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
            $escapedtargetou = getEscapedldapFilterStr $TargetMailUserOU
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
        $localusr = findLocalObject $localdc $srcObject
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
                    forceMergeObject $recipienttype $localOU $localusr $localDC $srcObject $srcDC
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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUDJ/s5dBgRGnRjKxRb+NHsfE1
# pxugghhqMIIE2jCCA8KgAwIBAgITMwAAASDzON/Hnq4y7AAAAAABIDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM4
# WhcNMjAwMTEwMjEwNzM4WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046MjI2NC1FMzNFLTc4MEMxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCO1OidLADhraPZx5FTVbd0PlB1xUfJ0J9zuRe1282yigKI
# +r7rvHTBllcSjV+E6G3BKO1FX7oV2CGaAGduTl2kk0vGSlrXC48bzR0SAb1Ui49r
# bUJTA++yfZA+34s8vYUye1XX2T5D0GKukK1hLkf8d7p2A5nygvMtnnybzmEVavSd
# g8lYzjK2EuekiLzL/lYUxAp2vRNFUitr7MHix5iU2nHEG4yU8crlXjYFgJ7q3CFv
# Il1yMsP/j+wk+1oCC1oLV6iOBcpq0Nxda/o+qN78nQFoQssfHoA9YdBGUnRHk+dK
# Sq5+GiV3AY0TRad2ZRzLcIcNmUJXny26YG+eokTpAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUIkw9WwdWW+zV8Il/Jq7A7bh6G7cwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAE4tuQuXzzaC2OIk4
# ZhJanhsgQv9Tk8ns/9elb8pAgYyZlSwxUtovV8Pd70jtAt0U/wjGd9n+QQJZKILM
# 6WCIieZFkZbqT9Ut9zA+tc2eQn4mt62PlyA+YJZNHEiPZhwgbjfLIwMRsm845B4N
# KN7WmfYwspHdT/mPgLWaBsSWS80PuAtpG3N+o9eTHskT+qauYAMqhZExfI8S2Rg4
# kdqAm7EU/Nroe4g0p+eKw6CAQ2ZuhuqHMMPgcQlSejcEbpS5WAzdCRd6qDXPHh0r
# C3FayhXrwu/KKuNW2hR1ZCx/ieNiR8+lWt1JxXgWAttgaRtR3VqGlL4aolg41UCo
# XfN1IjCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
# AQELBQAwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYG
# A1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xOTA1MDIy
# MTM3NDZaFw0yMDA1MDIyMTM3NDZaMHQxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xHjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJVaxoZpRx00HvFVw2Z19mJUGFgU
# ZyfwoyrGA0i85lY0f0lhAu6EeGYnlFYhLLWh7LfNO7GotuQcB2Zt5Tw0Uyjj0+/v
# UyAhL0gb8S2rA4fu6lqf6Uiro05zDl87o6z7XZHRDbwzMaf7fLsXaYoOeilW7SwS
# 5/LjneDHPXozxsDDj5Be6/v59H1bNEnYKlTrbBApiIVAx97DpWHl+4+heWg3eTr5
# CXPvOBxPhhGbHPHuMxWk/+68rqxlwHFDdaAH9aTJceDFpjX0gDMurZCI+JfZivKJ
# HkSxgGrfkE/tTXkOVm2lKzbAhhOSQMHGE8kgMmCjBm7kbKEd2quy3c6ORJECAwEA
# AaOCAX4wggF6MB8GA1UdJQQYMBYGCisGAQQBgjdMCAEGCCsGAQUFBwMDMB0GA1Ud
# DgQWBBRXghquSrnt6xqC7oVQFvbvRmKNzzBQBgNVHREESTBHpEUwQzEpMCcGA1UE
# CxMgTWljcm9zb2Z0IE9wZXJhdGlvbnMgUHVlcnRvIFJpY28xFjAUBgNVBAUTDTIz
# MDAxMis0NTQxMzUwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1ApUwVAYD
# VR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEFBQcBAQRV
# MFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNVHRMBAf8E
# AjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBaD4CtLgCersquiCyUhCegwdJdQ+v9Go4i
# Elf7fY5u5jcwW92VESVtKxInGtHL84IJl1Kx75/YCpD4X/ZpjAEOZRBt4wHyfSlg
# tmc4+J+p7vxEEfZ9Vmy9fHJ+LNse5tZahR81b8UmVmUtfAmYXcGgvwTanT0reFqD
# DP+i1wq1DX5Dj4No5hdaV6omslSycez1SItytUXSV4v9DVXluyGhvY5OVmrSrNJ2
# swMtZ2HKtQ7Gdn6iNntR1NjhWcK6iBtn1mz2zIluDtlRL1JWBiSjBGxa/mNXiVup
# MP60bgXOE7BxFDB1voDzOnY2d36ztV0K5gWwaAjjW5wPyjFV9wAyMX1hfk3aziaW
# 2SqdR7f+G1WufEooMDBJiWJq7HYvuArD5sPWQRn/mjMtGcneOMOSiZOs9y2iRj8p
# pnWq5vQ1SeY4of7fFQr+mVYkrwE5Bi5TuApgftjL1ZIo2U/ukqPqLjXv7c1r9+si
# eOcGQpEIn95hO8Ef6zmC57Ol9Ba1Ths2j+PxDDa+lND3Dt+WEfvxGbB3fX35hOaG
# /tNzENtaXK15qPhErbCTeljWhLPYk8Tk8242Z30aZ/qh49mDLsiL0ksurxKdQtXt
# v4g/RRdFj2r4Z1GMzYARfqaxm+88IigbRpgdC73BmwoQraOq9aLz/F1555Ij0U3o
# rXDihVAzgzCCBgcwggPvoAMCAQICCmEWaDQAAAAAABwwDQYJKoZIhvcNAQEFBQAw
# XzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29m
# dDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# MB4XDTA3MDQwMzEyNTMwOVoXDTIxMDQwMzEzMDMwOVowdzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAn6Fssd/b
# SJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u0afRTK10MCAR6wfVVJUVSZQbQpKumFww
# JtoAa+h7veyJBw/3DgSY8InMH8szJIed8vRnHCz8e+eIHernTqOhwSNTyo36Rc8J
# 0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShGLieshk9VUgzkAyz7apCQMG6H81kwnfp+
# 1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpHZhEnKWaol+TTBoFKovmEpxFHFAmCn4Tt
# VXj+AZodUAiFABAwRu233iNGu8QtVJ+vHnhBMXfMm987g5OhYQK1HQ2x/PebsgHO
# IktU//kFw8IgCwIDAQABo4IBqzCCAacwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4E
# FgQUIzT42VJGcArtQPt2+7MrsMM1sw8wCwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcV
# AQQDAgEAMIGYBgNVHSMEgZAwgY2AFA6sgmBAVieX5SUT/CrhClOVWeSkoWOkYTBf
# MRMwEQYKCZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0
# MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHmC
# EHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQu
# Y3JsMFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwEwYDVR0l
# BAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggIBABCXisNcA0Q23em0rXfb
# znlRTQGxLnRxW20ME6vOvnuPuC7UEqKMbWK4VwLLTiATUJndekDiV7uvWJoc4R0B
# hqy7ePKL0Ow7Ae7ivo8KBciNSOLwUxXdT6uS5OeNatWAweaU8gYvhQPpkSokInD7
# 9vzkeJkuDfcH4nC8GE6djmsKcpW4oTmcZy3FUQ7qYlw/FpiLID/iBxoy+cwxSnYx
# PStyC8jqcD3/hQoT38IKYY7w17gX606Lf8U1K16jv+u8fQtCe9RTciHuMMq7eGVc
# WwEXChQO0toUmPU8uWZYsy0v5/mFhsxRVuidcJRsrDlM1PZ5v6oYemIp76KbKTQG
# dxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOkqHQ1uE7FcSMSJnYLPFKMcVpGQxS8s7Ow
# TWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQQ+kr6bv0SMws1NgygEwmKkgkX1rqVu+m
# 3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFqPYmhdmG0bqETpr+qR/ASb/2KMmyy/t9R
# yIwjyWa9nR2HEmQCPS2vWY+45CHltbDKY7R4VAXUQS5QrJSwpXirs6CWdRrZkocT
# dSIvMqgIbqBbjCW/oO+EyiHW6x5PyZruSeD3AWVviQt9yGnI5m7qp5fOMSn/DsVb
# XNhNG6HY+i+ePy5VFmvJE6P9MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCBKgwggSkAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAFRno2PQHGjDkEAAAAAAVEwCQYFKw4DAhoFAKCB
# vDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUr63qZKU0IeAGmK2SzxyVgYM8L9cw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBACU2R9QyJNfEa18WmMZ6a4aOa2ZK2ng9qkArADsjBGiE
# XRuDSHO5eafZy6oX7rqt27oornKWajAjulvVfIV2GQS7Nngq8N1qs5HuzgDgUF/w
# AbZYZE4l5u4nkTZmBBfN/MfByqCHKz8wl+tgZJ/fqbxDir4orSpmqyyqTDtkolaK
# Eh65yo4rvwK66RcrqsXY+KYke58hsJL2Ok7DCJClj5VMRX115OHVdyDt/DTt0OFC
# hUO+J0cv0rcMkd74sWMt5/rQIr4dN+udKkHDX3XTqWcBhUzfdYvLQyEzF+rmO6gq
# JIhLaLPI5yItgkAZ4/6D2ES9tHchbIFOp07ptguBA2KhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# IPM438eerjLsAAAAAAEgMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDFaMCMGCSqGSIb3DQEJ
# BDEWBBQJ4UuZei/NG582HkejQ8xvWq17uTANBgkqhkiG9w0BAQUFAASCAQBmqY36
# hLyKqrMlgAhzVRex49pUZDbEuCNeL8D5fWM2FRbnr6BSgbxThgjHCK9Nq5THpVZX
# 0T13+NKySWGstfxFUcfrehNXNtsJzs/63oARb4kiMHvBpTqtHoSAEIp3xoUR1TD5
# Hyu0fecDihlLknwKotgcuPS+JLpF6ygQDC2ync8Il75e8hWCVRaiU21zYzeLfW25
# 438nKjaokpDxQ+FuA1LU0Jm8pikzfTdGgm+YeQA+TpZCBWsVIg8QTfDFHBBqUjfq
# zxoGclo2TtSTlbKv06ogT8bmPbVTt6PUzYNiBUMwN8EL3sfQux0wV2Hk3/LUpluD
# QJw2dFHqtHQ9MIMm
# SIG # End signature block
