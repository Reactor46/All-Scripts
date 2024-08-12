# .SYNOPSIS
# This script provides methods for import and export sync site mailbox script.
#
# Copyright (c) 2012 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# 
# .DESCRIPTION
# This script provides methods for import and export sync site mailbox script.
# It is used by Import-SyncSiteMailbox.ps1, Export-SyncSiteMailbox.ps1 and
# SyncSiteMailbox.ps1.
#

#
# Variables
#
# Office 365 to connect with
$ExchangeOnlineUrl = "https://pilot.outlook.com/PowerShell-LiveID/"

# Sync site mailbox cache file name
$LastSyncSiteMailboxCacheFile = "SyncSiteMailboxes.csv"

# Interval in days to check if a site mailbox has been deleted in Office 365
# Larger interval will gain better performance since deletion is not happening often
$DeletionCheckInternval = 10

# A cache to map Office 365 user's distinguished name with its primary smtp address
# on-premise cannot understand Office 365 user's distinguished name, it has be
# convert into primary smtp address
# Distingulish Name -> primary smtp address
$CloudUserDNToEmailsCache = @{}

# This format of accept domain is also coexistent domain
$CoexistentDomainsTemplate = ".mail.onmicrosoft.com"

# Switch to decide if syncing site mailbox properties
# Once DirSync client supports those 4 site mailbox specific properties,
# this should be false
$SyncSiteMailboxProperties = $true

# A cache to indicate if the email address exists in on-premise
# Email address -> $true or $false
$OnPremEmailsExistenceCache = @{}

######################################################################
# Common functions
######################################################################
#
# Run a cmdlet silently without throwing error
#
# Return:
# true if the cmdlet succeeds without error
# else, false
#
function RunCmdletSilently(
    [string] $cmdlet,
    [object] $parameters,
    [object] $pipelineObjects = $null,
    [ref] $output = [ref]$null)
{
    $parameters["ErrorAction"] = "Continue"
    $Global:Error.Clear()
    
    $result = $null
    try
    {
        if ($null -ne $pipelineObjects)
        {
            $result = ($pipelineObjects | &$cmdlet @parameters)
        }
        else
        {
            $result = (&$cmdlet @parameters)
        }
    
        if ($Global:Error.Count -eq 0)
        {
            $output.Value = $result
            return $true
        }
    }
    catch
    {
    }

    return $false
}

#
# Create a log file base on current date
#
# Return:
# $null if log file cannot be created
# else, log file name with full path
#
function CreateLogFile(
    [string] $logFileKey,
    [string] $folderPath)
{
    $date = [DateTime]::Now.ToString("yyyyMMdd")
    $logFileName = "$folderPath\$logFileKey-$date.txt"
        
    if ( -not (Test-Path -Path $logFileName))
    {
        $parameters = @{
            "Path" = $logFileName;
            "Type" = "File";
            "Force" = $true
        }

        if ( -not (RunCmdletSilently "New-Item" $parameters))
        {
            return $null
        }
    }

    return $logFileName
}

#
# Write a log entry
#
# Return:
# nothing
#
function WriteLog([string] $logData, [string]$category, [string]$logFile)
{
    $date = [DateTime]::Now.ToString("s")
    $fullLog = "$date, $category, $logData"
    
    $params = @{
        "FilePath" = $logFile;
        "Append" = $true;
        "InputObject" = $fullLog;
    }
    
    [void] (RunCmdletSilently "Out-File" $params)
} 

#
# Convert an object to string if it is array,
# "," is used as separator
#
# Return
# a string or $null
#
function ConverttoString($object)
{
    if ($null -eq $object)
    {
        return $null
    }
    
    return [string]::Join(",", [array]$object)
}

#
# Convert string to an array
# "," is used as separator
#
# Return
# array or $null
#
function ConverttoArray([string] $str)
{
    if ([string]::IsNullOrEmpty($str))
    {
        return $null
    }
    
    return $str.Split(",")
}

######################################################################
# Cache file support functions
######################################################################
#
# Load sync site mailboxes from cache file
#
# Return:
# a hashtable of sync site mailboxes with ExternalDirectoryObjectId as key if cache file is valid
# else, $null
#
function LoadSyncSiteMailboxesFromCache(
    [ref] $lastSyncUtcTime,
    [string] $folderPath
    )
{
    $syncSiteMailboxesCacheFile = "$folderPath\\$LastSyncSiteMailboxCacheFile"
    
    # the default value is the first site mailbox can be created in cloud
    $lastSyncUtcTime.Value = [DateTime]"2012-01-01"
    $syncSiteMailboxesArray = $null
    
    $syncSiteMailboxProperties = (
        "ExternalDirectoryObjectId", # ExternalDirectoryObjectId of the site mailbox in Office 365. It is used to identify an object in MSO and EXO.
        "WhenChangedUtc", # Utc date time: when site mailbox is changed in Office 365 in UTC format.
        "WhenCommittedUtc", # Utc date time: when this change is submitted into on-premise. That WhenCommittedUtc is less than WhenChangedUtc means we need a new commit.
        "InMSO", # true, false: is it synced into MSO. It can be committed only when it is in MSO.
        "IsDeleted", # true, false: is it deleted from EXO.
        "WhenDeletionChecked", # Utc date time: indicate when its deletion status is checked.
        "Name",
        "Alias",
        "DisplayName",
        "PrimarySmtpAddress",
        "EmailAddresses",
        "ExternalEmailAddress",
        "SharePointUrl",
        "SiteMailboxClosedTime",
        "SiteMailboxOwners",
        "SiteMailboxUsers",
        "LastCommitComment") # Error or information of last tried commit.
        
    $parameters = @{
        "Path" = $syncSiteMailboxesCacheFile;
        "Header" = $syncSiteMailboxProperties
    }
    
    if (Test-Path -Path $syncSiteMailboxesCacheFile)
    {
        if (RunCmdletSilently "Import-Csv" $parameters $null ([ref]$syncSiteMailboxesArray))
        {
            $syncSiteMailboxesHashtable = @{}
            $syncSiteMailboxesArray = [array]$syncSiteMailboxesArray
            if ($null -eq $syncSiteMailboxesArray)
            {
                $syncSiteMailboxesArray = @()
            }
            
            foreach ($syncSiteMailbox in $syncSiteMailboxesArray)
            {
                if(IsValidSyncSiteMailbox $syncSiteMailbox)
                {
                    $syncSiteMailboxesHashtable[$syncSiteMailbox.ExternalDirectoryObjectId] = $syncSiteMailbox
                }
                
                # The last sync Utc time should be the latest one
                $syncSiteMailboxWhenChangedUtc = [DateTime]::MinValue
                if ([DateTime]::TryParse($syncSiteMailbox.WhenChangedUtc, ([ref]$syncSiteMailboxWhenChangedUtc)) -and
                    $lastSyncUtcTime.Value -lt $syncSiteMailboxWhenChangedUtc)
                {
                    $lastSyncUtcTime.Value = $syncSiteMailboxWhenChangedUtc
                }
            }
            
            return $syncSiteMailboxesHashtable
        }
    }
    
    return $null
}

#
# Create a sync site mailbox PSObject
#
# Return:
# A custom PSOjbect holding sync site mailbox properties
# all properties are of string type for saving or loading from csv file
# 
function CreateSyncSiteMailboxPSObject(
        $externalDirectoryObjectId,
        $whenChangedUtc,
        $whenCommittedUtc,
        $inMSO,
        $isDeleted,
        $whenDeletionChecked,
        $name,
        $alias,
        $displayName,
        $primarySmtpAddress,
        $emailAddresses,
        $externalEmailAddress,
        $sharePointUrl,
        $siteMailboxClosedTime,
        $siteMailboxOwners,
        $siteMailboxUsers,
        $lastCommitComment)
{
    $psObject = New-Object -TypeName PSObject
    
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "ExternalDirectoryObjectId" -Value (ConverttoString $externalDirectoryObjectId)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "WhenChangedUtc" -Value (ConverttoString $whenChangedUtc)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "WhenCommittedUtc" -Value (ConverttoString $whenCommittedUtc)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "InMSO" -Value (ConverttoString $inMSO)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "IsDeleted" -Value (ConverttoString $isDeleted)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "WhenDeletionChecked" -Value (ConverttoString $whenDeletionChecked)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "Name" -Value (ConverttoString $name)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "Alias" -Value (ConverttoString $alias)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "DisplayName" -Value (ConverttoString $displayName)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "PrimarySmtpAddress" -Value (ConverttoString $primarySmtpAddress)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "EmailAddresses" -Value (ConverttoString $emailAddresses)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "ExternalEmailAddress" -Value (ConverttoString $externalEmailAddress)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "SharePointUrl" -Value (ConverttoString $sharePointUrl)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "SiteMailboxClosedTime" -Value (ConverttoString $siteMailboxClosedTime)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "SiteMailboxOwners" -Value (ConverttoString $siteMailboxOwners)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "SiteMailboxUsers" -Value (ConverttoString $siteMailboxUsers)
    Add-Member -InputObject $psObject -MemberType:NoteProperty -Name "LastCommitComment" -Value (ConverttoString $lastCommitComment)

    return $psObject
}

#
# Verify if a sync site mailbox PSObject is valid
#
# Return:
# true if it is valid
# else, false
#
function IsValidSyncSiteMailbox(
    [object] $syncSiteMailbox)
{
    if ($null -eq $syncSiteMailbox)
    {
        return $false
    }
    
    # Ensure ExternalDirectoryObjectId is GUID
    if ( -not [Guid]::TryParse($syncSiteMailbox.ExternalDirectoryObjectId, [ref][Guid]::Empty))
    {
        return $false
    }
    
    # Ensure WhenChangedUtc is DateTime
    if ( -not [DateTime]::TryParse($syncSiteMailbox.WhenChangedUtc, [ref][DateTime]::UtcNow))
    {
        return $false
    }
    
    # Ensure WhenCommittedUtc is null or valid DateTime
    if (( -not [string]::IsNullOrEmpty($syncSiteMailbox.WhenCommittedUtc)) -and
        ( -not [DateTime]::TryParse($syncSiteMailbox.WhenChangedUtc, [ref][DateTime]::UtcNow)))
    {
        return $false
    }

    # Ensure WhenDeletionChecked is null or valid DateTime
    if (( -not [string]::IsNullOrEmpty($syncSiteMailbox.WhenDeletionChecked)) -and
        ( -not [DateTime]::TryParse($syncSiteMailbox.WhenDeletionChecked, [ref][DateTime]::UtcNow)))
    {
        return $false
    }
    
    # Ensure Name, PrimarySmtpAddress, EmailAddresses, ExternalEmailAddress have value
    if ([string]::IsNullOrEmpty($syncSiteMailbox.Name) -or
        [string]::IsNullOrEmpty($syncSiteMailbox.PrimarySmtpAddress) -or
        [string]::IsNullOrEmpty($syncSiteMailbox.EmailAddresses) -or
        [string]::IsNullOrEmpty($syncSiteMailbox.ExternalEmailAddress))
    {
        return $false
    }
    
    return $true
}

#
# Save sync site mailboxes into cache file
#
# Return:
# true if the save succeeds
# else, false
#
function SaveSyncSiteMailboxesIntoCache(
    [Hashtable] $syncSiteMailboxesHashtable,
    [DateTime] $syncUtcTime,
    [string] $folderPath
    )
{
    $syncSiteMailboxesCacheFile = "$folderPath\\$LastSyncSiteMailboxCacheFile"
    
    $arrayList = New-Object -Type System.Collections.ArrayList
    foreach($syncSiteMailbox in $syncSiteMailboxesHashtable.Values)
    {
        [void] $arrayList.Add($syncSiteMailbox)
    }
    
    # build an cache entry to record last synced time
    $cacheEntryForSyncUtcTime = CreateSyncSiteMailboxPSObject -externalDirectoryObjectId ([guid]::Empty) -whenChangedUtc $syncUtcTime -lastCommitComment "This entry indicates the last site maiblox query Utc time."
    [void] $arrayList.Add($cacheEntryForSyncUtcTime)
    
    $parameters = @{
        "Path" = $syncSiteMailboxesCacheFile;
        "Force" = $true;
        "NoTypeInformation" = $true
    }
    
    if (RunCmdletSilently "Export-Csv" $parameters $arrayList)
    {
        return $true
    }
    
    return $false
}

######################################################################
# Office 365 functions
######################################################################
#
# Connect to Office 365
#
# Return:
# session instance if connection is established
# else, $null
#
function ConnecttoExchangeOnline(
    [System.Management.Automation.PSCredential] $credential)
{
    $session = $null
    
    $parameters = @{
        "ConfigurationName" = "Microsoft.Exchange";
        "ConnectionURI" = $ExchangeOnlineUrl
        "AllowRedirection" = $true;
        "SessionOption" = (New-PSSessionOption -SkipCACheck -SkipCNCheck);
        "Credential" = $credential;
        "Authentication" = "Basic";
    }
        
    if (RunCmdletSilently "New-PSSession" $parameters $null ([ref]$session))
    {
        return $session
    }

    return $null
}

#
# Load coexistent domain
#
# Return:
# An array of coexistent domains if they exist
# else, $null
#
function LoadCoexistentDomain(
    [string] $cmdlet)
{
    $fallbackCoexistentDomains = $null
    $coexistentDomains = New-Object -Type System.Collections.ArrayList
    $domains = $null
    if (RunCmdletSilently $cmdlet (@{}) $null ([ref]$domains))
    {
        if ($null -ne $domains)
        {
            foreach ($domain in $domains)
            {
                $domainString = $domain.DomainName.ToString()
                if ($domain.IsCoexistenceDomain -or
                    (-not [string]::IsNullOrEmpty($CoexistentDomainsTemplate) -and
                    $domainString.EndsWith($CoexistentDomainsTemplate, 1)))
                {
                    [void] $coexistentDomains.Add($domainString)
                }
            }
        }
    }
    
    return $coexistentDomains
}

#
# Calculate external email address for site mailbox
#
# Return:
# external email address if it exists
# else $null
#
function CalculateExternalEmailAddress(
    [array] $emailAddresses,
    [array] $coexistentDomains
)
{
    $externalEmail = $null

    if ($null -ne $emailAddresses -and $null -ne $coexistentDomains)
    {
        foreach ($email in $emailAddresses)
        {
            $emailString = $email.ToString()
            foreach ($domain in $coexistentDomains)
            {
                if ($emailString.EndsWith("@$domain", 1))
                {
                    $externalEmail = $emailString

                    if ([string]::IsNullOrEmpty($CoexistentDomainsTemplate) -or
                        $emailString.EndsWith("@$CoexistentDomainsTemplate", 1))
                    {
                        return $externalEmail
                    }
                }
            }
        }
    }

    return $externalEmail
}

#
# Fill site mailbox properties into the sync site mailbox PSObject
#
# Return:
# true, if it is filled successfully
# else, false
#
function FillSiteMailboxProperties(
    [object]$syncSiteMailbox,
    [string] $getSiteMailboxCmdlet,
    [string] $getRecipientCmdlet)
{
    if ( -not $SyncSiteMailboxProperties)
    {
        return $false
    }

    # load site mailbox
    $siteMailboxProps = $null
    $params = @{
        "Identity" = $primarySmtpAddress;
        "BypassOwnerCheck" = $true
    }
    
    if (RunCmdletSilently $getSiteMailboxCmdlet $params $null ([ref]$siteMailboxProps))
    {
        if ($null -ne $siteMailboxProps -and ([array]$siteMailboxProps).Count -eq 1)
        {
            $syncSiteMailbox.SharePointUrl = (ConverttoString $siteMailboxProps.SharePointUrl)
            $syncSiteMailbox.SiteMailboxClosedTime = (ConverttoString $siteMailboxProps.ClosedTime)
            $syncSiteMailbox.SiteMailboxOwners = (ConverttoString (ConvertDNsToPrimarySmtpAddresses $siteMailboxProps.Owners $getRecipientCmdlet))
            
            $siteMailboxUsers = New-Object -Type System.Collections.ArrayList
            if ($null -ne $siteMailboxProps.Owners)
            {
                [void] $siteMailboxUsers.AddRange($siteMailboxProps.Owners)
            }
            
            if ($null -ne $siteMailboxProps.Members)
            {
                [void] $siteMailboxUsers.AddRange($siteMailboxProps.Members)
            }
            
            $syncSiteMailbox.SiteMailboxUsers = (ConverttoString (ConvertDNsToPrimarySmtpAddresses $siteMailboxUsers $getRecipientCmdlet))
            return $true
        }
    }
    else
    {
        # Check if there is an explicit exception to indicate the object is not found:
        # Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException
        if ($Error.Count -gt 0 -and
            $Error[0].Exception.SerializedRemoteException.ToString().Contains("Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException"))
        {
            $syncSiteMailbox.InMSO = $true.ToString()
            return $true
        }
    }
}

#
# Filter out non-smtp address from email address list
#
# Return:
# a list of smtp address
# or empty
#
function FilterOutNonSmtpAddress([array] $emailAddresses)
{
    $smtpAddresses = New-Object -Type System.Collections.ArrayList

    if ($null -ne $emailAddresses)
    {
        foreach ($addr in $emailAddresses)
        {
            $addrString = $addr.ToString()
            if ($addrString.StartsWith("smtp",1))
            {
                [void] $smtpAddresses.Add($addr)
            }
        }
    }

    return $smtpAddresses
}

#
# Remove all smtp addresses from first parameter and add smtp address into the list 
#
# Return:
# a list of addresses in string type
# or empty
#
function ReplaceSmtpAddresses([array] $emailAddresses, [array] $smtpAddresses)
{
    $replacedEmailAddresses = New-Object -Type System.Collections.ArrayList

    if ($null -ne $emailAddresses)
    {
        foreach( $addr in $emailAddresses)
        {
            $addrString = $addr.ToString()
            if (-not $addrString.StartsWith("smtp",1))
            {
                [void] $replacedEmailAddresses.Add($addrString)
            }
        }
    }

    if ($null -ne $smtpAddresses)
    {
        foreach ($addr in $smtpAddresses)
        {
            $addrString = $addr.ToString()
            [void] $replacedEmailAddresses.Add($addrString)
        }
    }

    return $replacedEmailAddresses
}

#
# Merge a site mailbox into sync site mailboxes hashtable
#
# Return:
# true if it is merged successfully
# else false
#
function MergeSiteMailboxIntoCache(
    [object] $recipient,
    [hashtable] $syncSiteMailboxesHashtable,
    [array] $coexistentDomains,
    [string] $getSiteMailboxCmdlet,
    [string] $getRecipientCmdlet)
{
    $smtpAddresses = @(FilterOutNonSmtpAddress $recipient.EmailAddresses)

    $externalDirectoryObjectId = $recipient.ExternalDirectoryObjectId.ToString()
    $whenChangedUtc = (ConverttoString $recipient.WhenChangedUtc)
    $whenCommittedUtc = $null
    $inMSO = $false
    $isDeleted = $false
    $whenDeletionChecked = $whenChangedUtc
    $name = $recipient.Name
    $alias = $recipient.Alias
    $displayName = $recipient.DisplayName
    $primarySmtpAddress = (ConverttoString $recipient.PrimarySmtpAddress)
    $emailAddresses = (ConverttoString $smtpAddresses)
    $externalEmailAddress = (CalculateExternalEmailAddress $smtpAddresses $coexistentDomains)
   
    $sharePointUrl = $null
    $siteMailboxClosedTime = $null
    $siteMailboxOwners = $null
    $siteMailboxUsers = $null
    $lastCommitComment = $null
    
    $newSyncSiteMailbox = CreateSyncSiteMailboxPSObject `
        $externalDirectoryObjectId `
        $whenChangedUtc `
        $whenCommittedUtc `
        $inMSO `
        $isDeleted `
        $whenDeletionChecked `
        $name `
        $alias `
        $displayName `
        $primarySmtpAddress `
        $emailAddresses `
        $externalEmailAddress `
        $sharePointUrl `
        $siteMailboxClosedTime `
        $siteMailboxOwners `
        $siteMailboxUsers `
        $lastCommitComment

    # Load site mailbox specific properties
    if ( -not (FillSiteMailboxProperties $newSyncSiteMailbox $getSiteMailboxCmdlet $getRecipientCmdlet))
    {
        $newSyncSiteMailbox.LastCommitComment = "Cannot load site mailbox properties."
    }

    # Merge into the cache file
    if ($syncSiteMailboxesHashtable.ContainsKey($externalDirectoryObjectId))
    {
        $syncSiteMailbox = $syncSiteMailboxesHashtable[$externalDirectoryObjectId]
        
        # Update WhenCommittedUtc to WhenChangedUtc to avoid unnecessary export
        # if base properties are not actually changed and there is no pending change
        if ( $newSyncSiteMailbox.Alias -eq $syncSiteMailbox.Alias -and
             $newSyncSiteMailbox.DisplayName -eq $syncSiteMailbox.DisplayName -and
             $newSyncSiteMailbox.PrimarySmtpAddress -eq $syncSiteMailbox.PrimarySmtpAddress -and
             $newSyncSiteMailbox.EmailAddresses -eq $syncSiteMailbox.EmailAddresses -and
             $newSyncSiteMailbox.ExternalEmailAddress -eq $syncSiteMailbox.ExternalEmailAddress -and
             $newSyncSiteMailbox.SharePointUrl -eq $syncSiteMailbox.SharePointUrl -and
             $newSyncSiteMailbox.SiteMailboxClosedTime -eq $syncSiteMailbox.SiteMailboxClosedTime -and
             $newSyncSiteMailbox.SiteMailboxOwners -eq $syncSiteMailbox.SiteMailboxOwners -and
             $newSyncSiteMailbox.SiteMailboxUsers -eq $syncSiteMailbox.SiteMailboxUsers -and
             $syncSiteMailbox.WhenChangedUtc -le $syncSiteMailbox.WhenCommittedUtc
             )
        {
            $newSyncSiteMailbox.whenCommittedUtc = $newSyncSiteMailbox.WhenChangedUtc
        }
        
        $newSyncSiteMailbox.InMSO = $syncSiteMailbox.InMSO
        $newSyncSiteMailbox.WhenDeletionChecked = $syncSiteMailbox.WhenDeletionChecked
    }
    
    if (IsValidSyncSiteMailbox $newSyncSiteMailbox)
    {
        $syncSiteMailboxesHashtable[$externalDirectoryObjectId] = $newSyncSiteMailbox
        return $true
    }
    
    return $false
}

#
# Check if a site mailbox is deleted from Office 365
#
# Return:
# true, if it has been deleted from Office 365
# else, false
#
function CheckIfSiteMailboxDeletedInEXO(
    [object] $syncSiteMailbox,
    [string] $cmdlet
    )
{
    if ("True" -eq $syncSiteMailbox.IsDeleted)
    {
        return $true
    }
    
    if ([string]::IsNullOrEmpty($syncSiteMailbox.WhenDeletionChecked) -or
        ([DateTime]$syncSiteMailbox.WhenDeletionChecked).AddDays($DeletionCheckInternval) -lt [DateTime]::UtcNow)
    {
        $syncSiteMailbox.WhenDeletionChecked = [DateTime]::UtcNow.ToString()
        
        # Only perform actual check every $DeletionCheckInternval days
        $parameters = @{
            "Identity" = $syncSiteMailbox.ExternalDirectoryObjectId;
            "BypassOwnerCheck" = $true
        }
        
        if (RunCmdletSilently $cmdlet $parameters)
        {
            return $false
        }
        
        # Check if there is an explicit exception to indicate the object is not found:
        # Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException
        if ($Error.Count -gt 0 -and
            $Error[0].Exception.SerializedRemoteException.ToString().Contains("Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException"))
        {
            $syncSiteMailbox.IsDeleted = $true.ToString()
            $syncSiteMailbox.WhenChangedUtc = $syncSiteMailbox.WhenDeletionChecked
            return $true
        }
    }
    
    return $false
}

#
# Resolve DN to primary smtp address
#
# Return:
# distinguished name or $null
#
function ResolveDNToPrimarySmtpAddress(
    [string] $identity,
    [string] $cmdlet
)
{
    if ( -not $CloudUserDNToEmailsCache.ContainsKey($identity))
    {
        $CloudUserDNToEmailsCache[$identity] = $null
    
        $recipient = $null
        $params = @{
            "Identity" = $identity
        }
    
        if (RunCmdletSilently $cmdlet $params $null ([ref]$recipient))
        {
           if ($null -ne $recipient -and ([array]$recipient).Count -eq 1)
           {
                $CloudUserDNToEmailsCache[$identity] = $recipient.PrimarySmtpAddress.ToString()
           }
        }
    }
    
    return $CloudUserDNToEmailsCache[$identity]
}

#
# Resolve a list of DNs to a list primary smtp addresses
#
# Return:
# a list of primary smtp addresses or $null
#
function ConvertDNsToPrimarySmtpAddresses(
    [array] $distinguishedNames,
    [string] $cmdlet)
{
    $primarySmtpAddresses = New-Object -Type System.Collections.ArrayList
    if ($null -ne $distinguishedNames)
    {
        foreach ($dn in $distinguishedNames)
        {
            $primarySmtpAddress = ResolveDNToPrimarySmtpAddress $dn $cmdlet
            if ( -not [string]::IsNullOrEmpty($primarySmtpAddress))
            {
                [void] $primarySmtpAddresses.Add($primarySmtpAddress)
            }
        }
    }

    return $primarySmtpAddresses
}

######################################################################
# MSO online functions
######################################################################
#
# Check if a site mailbox is in MSO online
#
# Return:
# true, if the object is getted from Microsoft Online Service
# else, false
#
function CheckIfSiteMailboxInMSO(
    [string] $id,
    [string] $cmdlet)
{
    $parameters = @{
        "ObjectId" = $id
    }
    
    if (RunCmdletSilently $cmdlet $parameters)
    {
        return $true
    }
    
    return $false
}

######################################################################
# On-Premise local PowerShell functions
######################################################################
#
# Check validness of site mailbox users
#
# Return:
# A valid array of site mailbox users
#
function ValidateSiteMailboxUsers(
    [array] $usersEmail,
    [string] $cmdlet)
{
    $validUsersEmail = New-Object -Type System.Collections.ArrayList
    
    $parameters = @{
        "Identity" = $null;
        "RecipientTypeDetails" = @("UserMailbox","LinkedMailbox","RemoteUserMailbox","MailUser")
    }

    if ($null -ne $usersEmail)
    {
        foreach ($userEmail in ([array]$usersEmail))
        {
            if ( -not $OnPremEmailsExistenceCache.ContainsKey($userEmail))
            {
                $parameters["Identity"] = $userEmail
                $OnPremEmailsExistenceCache[$userEmail] = RunCmdletSilently $cmdlet $parameters
            }
        
            if ($OnPremEmailsExistenceCache[$userEmail] -and
                (-not $validUsersEmail.Contains($userEmail)))
            {
                [void] $validUsersEmail.Add($userEmail)
            }
        }
    }
    
    return $validUsersEmail
}

#
# Update sync site mailbox in on-premise
#
# Return:
# update status string
#
function UpdateOnPremSyncSiteMailbox(
    [object] $syncSiteMailbox,
    [string] $getCmdlet,
    [string] $setCmdlet,
    [string] $newCmdlet,
    [string] $delCmdlet)
{
    # Not ready in MSO
    if ("True" -ne $syncSiteMailbox.InMSO)
    {
        return "skip since it is not ready in Microsoft Online Service."
    }

    # No changes since last change
    if (( -not [string]::IsNullOrEmpty($syncSiteMailbox.WhenCommittedUtc)) -and
        ([DateTime]$syncSiteMailbox.WhenChangedUtc -le [DateTime]$syncSiteMailbox.WhenCommittedUtc))
    {
        return "skip since there is no new changes."
    }

    #
    # Handle deletion first
    #
    $utcNow = [DateTime]::UtcNow
    if ( "True" -eq $syncSiteMailbox.IsDeleted)
    {
        $parameters = @{
            "Identity" = $syncSiteMailbox.PrimarySmtpAddress;
            "Confirm" = $false
        }
        
        if (RunCmdletSilently $delCmdlet $parameters)
        {
            $syncSiteMailbox.WhenCommittedUtc = $syncSiteMailbox.WhenChangedUtc
            $syncSiteMailbox.LastCommitComment = "Delete on $utcNow by $delCmdlet."
            
            return "delete successfully."
        }

        $syncSiteMailbox.LastCommitComment = "Cannot be deleted by $delCmdlet on $utcNow because of $Error."
        return "fail to delete becase of $Error."
    }

    # Parameters for update and create
    $emailAddresses = ConverttoArray $syncSiteMailbox.EmailAddresses
    $updateCmdlet = $null
    $updateParameters = @{
        "Alias" = $syncSiteMailbox.Alias;
        "DisplayName" = $syncSiteMailbox.DisplayName;
        "ExternalEmailAddress" = $syncSiteMailbox.ExternalEmailAddress;
        "RecipientDisplayType" = "SyncedTeamMailboxUser"
    }

    if ($SyncSiteMailboxProperties)
    {
        $updateParameters["SharePointUrl"] = $null
        if (-not [string]::IsNullOrEmpty($syncSiteMailbox.SharePointUrl))
        {
            $updateParameters["SharePointUrl"] = $syncSiteMailbox.SharePointUrl
        }
        
        $updateParameters["SiteMailboxClosedTime"] = $null
        if (-not [string]::IsNullOrEmpty($syncSiteMailbox.SiteMailboxClosedTime))
        {
            $updateParameters["SiteMailboxClosedTime"] = $syncSiteMailbox.SiteMailboxClosedTime
        }

        $updateParameters["SiteMailboxOwners"] = ValidateSiteMailboxUsers (ConverttoArray $syncSiteMailbox.SiteMailboxOwners) $getCmdlet
        $updateParameters["SiteMailboxUsers"] = ValidateSiteMailboxUsers (ConverttoArray $syncSiteMailbox.SiteMailboxUsers) $getCmdlet
    }

    # Load existing sync site mailbox from active directory
    $siteMailboxProps = $null
    $parameters = @{
        "Identity" = $syncSiteMailbox.PrimarySmtpAddress
    }

    [void](RunCmdletSilently $getCmdlet $parameters $null ([ref]$siteMailboxProps))
    if ($null -ne $siteMailboxProps -and ([array]$siteMailboxProps).Count -eq 1)
    {
        # Handle update
        $updateCmdlet = $setCmdlet
        $updateParameters["Identity"] = $syncSiteMailbox.PrimarySmtpAddress
        $updateParameters["EmailAddresses"] = ReplaceSmtpAddresses ([array]$siteMailboxProps.EmailAddresses) ([array]$emailAddresses)
    }
    else
    {
        # Hanlde create
        $updateCmdlet = $newCmdlet
        $updateParameters["Name"] = $syncSiteMailbox.Name
        $updateParameters["EmailAddresses"] = $emailAddresses

        if ($updateParameters.ContainsKey("SiteMailboxOwners") -and
            $null -eq $updateParameters["SiteMailboxOwners"])
        {
            $updateParameters.Remove("SiteMailboxOwners")
        }

        if ($updateParameters.ContainsKey("SiteMailboxUsers") -and
            $null -eq $updateParameters["SiteMailboxUsers"])
        {
            $updateParameters.Remove("SiteMailboxUsers")
        }
    }

    $now = [DateTime]::Now.ToString("s")
    if (RunCmdletSilently $updateCmdlet $updateParameters)
    {
        $syncSiteMailbox.WhenCommittedUtc = $syncSiteMailbox.WhenChangedUtc
        $syncSiteMailbox.LastCommitComment = "Update on $now by $updateCmdlet."
    }
    else
    {
        $syncSiteMailbox.LastCommitComment = "Cannot be updated on $now by $updateCmdlet because of $Error."
    }

    return $syncSiteMailbox.LastCommitComment
}
# SIG # Begin signature block
# MIIdtAYJKoZIhvcNAQcCoIIdpTCCHaECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUrioqTjQWM8DGh743mYqbfoLb
# rtmgghhkMIIEwzCCA6ugAwIBAgITMwAAAK7sP622i7kt0gAAAAAArjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwNTAzMTcxMzI1
# WhcNMTcwODAzMTcxMzI1WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkI4RUMtMzBBNC03MTQ0MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAxTU0qRx3sqZg
# 8GN4YCrqA1CzmYPp8+U/MG7axHXPZGdMvNbRSPl29ba88jCYRut/6p5OjvCGNcRI
# MPWKFMqKVeY8zUoQNp46jYsXenl4vTAgJ2cUCeaGy9vxLYTGuXtaChn+jIpPuR6x
# UQ60Y44M2jypsbcQZYc6Oukw4co+CIw8fKqxPcDjdm1c/gyzVnhSYTXsv8S0NBwl
# iuhNCNE4D8b0LNj7Exj5zfVYGvP6Z+JtGY7LT+7caUCT0uItKlE0D/iDvlY5zLrb
# luUb4WLUBpglMw7bU0BSAcvcNx0XyV7+AdcmhiFQGt4pZjbVzOsXs3POWHTq4/KX
# RmtGHKfvMwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFBw4ctJakrpBibpB9TJkYJsJ
# gGBUMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAAAZsVbJVNFZUNMcXRxKeelc1DgiQHLC60Sika98OwDFXomY
# akk6yvE+fJ3DICnDUK9kmf83sYTOQ5Y7h3QzwHcPdyhLPHSBBmuPklj6jcWGuvHK
# pUuP9PTjyKBw0CPZ1PTO1Jc5RjsQYvxqu01+G5UvZolnM6Ww7QpmBoDEyze5J+dg
# GwrWMhIKDzKLV9do6R5ouZQvLvV7bjH50AX2tK2n3zpZYvAl/LayLHFNIO7A2DQ1
# VzWa3n2yyYvameaX1NkSLA32PqjAXykmkDfHQ6DFVuDV4nqrNI+s14EJgMQy8DzU
# 9X7+KIkCzLFNq/bc2WDo15qsQiACPVSKY1IOGiIwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBLowggS2AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBzjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUcRA5Zp/njJAN17nvCxh24+TxhHswbgYKKwYB
# BAGCNwIBDDFgMF6gNoA0AFMAeQBuAGMAUwBpAHQAZQBNAGEAaQBsAGIAbwB4AEwA
# aQBiAHIAYQByAHkALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20v
# ZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAFY93SZVRCI8jGSFvOUoJ+aYBsFp
# +cTs9K8uln9TsFTF8YMNsZsGRpOHpWnZPHdfFTcOnhmuF89bgjcvuUji8rEeXyTi
# tnH5Pr/S8pqEHR7TYuGuKM4/nVRUVQQ7Mmv8fdso5Jy4OOtXJS5Dh2+zpH/GwRsT
# hQEJL1ERhdECYgM+aaIcSrpmoBrn/7haBhzI09kCHJ2Lro04JrGK24ltJ2B5riAq
# y5Mq4fzmWNsDywSfgZXETp8kUfkkFLspsB2l3WGa7in+znUu/rXj1vHoUlmOe/Vn
# VcXIc6jQOffM1qNwDllAE0vWFJEhtOhdr4PRO8wc9Mq1q+kLg9ic9BFXdjihggIo
# MIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBAhMzAAAAruw/rbaLuS3SAAAAAACuMAkGBSsOAwIaBQCgXTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ0MzFa
# MCMGCSqGSIb3DQEJBDEWBBRfPHWPvq/gm4IYhgUebKqV1X7eTDANBgkqhkiG9w0B
# AQUFAASCAQBV+zloqN4KlpYzqiyVua8HvJIkHgn6q5i5xZ/sdqNXkg7xA2E15UM3
# JlLebuecJgnHhbSvFxVRdVF/Z9eM48pwIr3DhKLT3PxfdV1z7oEXXcQO5Qg59N0Y
# 4gyJsTLWxh7BVlvoloQOyRim8QNCINNuCKM9Dbd19siQBpGemXmks0VYGXo4jgZL
# pgKs90pzQy28evUXVgoKXyNVSbeGtPZH5svVNgGUeUq/FCg/zxdDkccGaIFHExac
# 5guNsA5zSgQYSDp9roX/wrkaj7rq/UQ0OpjJjv4zX2xm746VLgPGs4BRbeYiDRit
# /e7QUNpeJL7TKmtLjxKNpCYvepBTpjLV
# SIG # End signature block
