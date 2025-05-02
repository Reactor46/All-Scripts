# .SYNOPSIS
#    Syncs mail-enabled public folder objects from the local Exchange deployment into O365. It uses the local Exchange deployment
#    as master to determine what changes need to be applied to O365. The script will create, update or delete mail-enabled public
#    folder objects on O365 Active Directory when appropriate.
#
# .DESCRIPTION
#    The script must be executed from an Exchange 2007 or 2010 Management Shell window providing access to mail public folders in
#    the local Exchange deployment. Then, using the credentials provided, the script will create a session against Exchange Online,
#    which will be used to manipulate O365 Active Directory objects remotely.
#
#    Copyright (c) 2014 Microsoft Corporation. All rights reserved.
#
#    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# .PARAMETER Credential
#    Exchange Online user name and password.
#
# .PARAMETER CsvSummaryFile
#    The file path where sync operations and errors will be logged in a CSV format.
#
# .PARAMETER ConnectionUri
#    The Exchange Online remote PowerShell connection uri. If you are an Office 365 operated by 21Vianet customer in China, use "https://partner.outlook.cn/PowerShell".
#
# .PARAMETER Confirm
#    The Confirm switch causes the script to pause processing and requires you to acknowledge what the script will do before processing continues. You don't have to specify
#    a value with the Confirm switch.
#
# .PARAMETER Force
#    Force the script execution and bypass validation warnings.
#
# .PARAMETER WhatIf
#    The WhatIf switch instructs the script to simulate the actions that it would take on the object. By using the WhatIf switch, you can view what changes would occur
#    without having to apply any of those changes. You don't have to specify a value with the WhatIf switch.
#
# .EXAMPLE
#    .\Sync-MailPublicFolders.ps1 -Credential (Get-Credential) -CsvSummaryFile:sync_summary.csv
#    
#    This example shows how to sync mail-public folders from your local deployment to Exchange Online. Note that the script outputs a CSV file listing all operations executed, and possibly errors encountered, during sync.
#
# .EXAMPLE
#    .\Sync-MailPublicFolders.ps1 -Credential (Get-Credential) -CsvSummaryFile:sync_summary.csv -ConnectionUri:"https://partner.outlook.cn/PowerShell"
#    
#    This example shows how to use a different URI to connect to Exchange Online and sync mail-public folders from your local deployment.
#
param(
    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential] $Credential,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string] $CsvSummaryFile,
    
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID",

    [Parameter(Mandatory=$false)]
    [bool] $Confirm = $true,

    [Parameter(Mandatory=$false)]
    [switch] $Force = $false,

    [Parameter(Mandatory=$false)]
    [switch] $WhatIf = $false
)

# Writes a dated information message to console
function WriteInfoMessage()
{
    param ($message)
    Write-Host "[$($(Get-Date).ToString())]" $message;
}

# Writes a dated warning message to console
function WriteWarningMessage()
{
    param ($message)
    Write-Warning ("[{0}] {1}" -f (Get-Date),$message);
}

# Writes a verbose message to console
function WriteVerboseMessage()
{
    param ($message)
    Write-Host "[VERBOSE] $message" -ForegroundColor Green -BackgroundColor Black;
}

# Writes an error importing a mail public folder to the CSV summary
function WriteErrorSummary()
{
    param ($folder, $operation, $errorMessage, $commandtext)

    WriteOperationSummary $folder.Guid $operation $errorMessage $commandtext;
    $script:errorsEncountered++;
}

# Writes the operation executed and its result to the output CSV
function WriteOperationSummary()
{
    param ($folder, $operation, $result, $commandtext)

    $columns = @(
        (Get-Date).ToString(),
        $folder.Guid,
        $operation,
        (EscapeCsvColumn $result),
        (EscapeCsvColumn $commandtext)
    );

    Add-Content $CsvSummaryFile -Value ("{0},{1},{2},{3},{4}" -f $columns);
}

#Escapes a column value based on RFC 4180 (http://tools.ietf.org/html/rfc4180)
function EscapeCsvColumn()
{
    param ([string]$text)

    if ($text -eq $null)
    {
        return $text;
    }

    $hasSpecial = $false;
    for ($i=0; $i -lt $text.Length; $i++)
    {
        $c = $text[$i];
        if ($c -eq $script:csvEscapeChar -or
            $c -eq $script:csvFieldDelimiter -or
            $script:csvSpecialChars -contains $c)
        {
            $hasSpecial = $true;
            break;
        }
    }

    if (-not $hasSpecial)
    {
        return $text;
    }
    
    $ch = $script:csvEscapeChar.ToString([System.Globalization.CultureInfo]::InvariantCulture);
    return $ch + $text.Replace($ch, $ch + $ch) + $ch;
}

# Writes the current progress
function WriteProgress()
{
    param($statusFormat, $statusProcessed, $statusTotal)
    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity `
        -Status ($statusFormat -f $statusProcessed,$statusTotal) `
        -PercentComplete (100 * ($script:itemsProcessed + $statusProcessed)/$script:totalItems);
}

# Create a tenant PSSession against Exchange Online.
function InitializeExchangeOnlineRemoteSession()
{
    WriteInfoMessage $LocalizedStrings.CreatingRemoteSession;

    $oldWarningPreference = $WarningPreference;
    $oldVerbosePreference = $VerbosePreference;

    try
    {
        $VerbosePreference = $WarningPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue;
        $sessionOption = (New-PSSessionOption -SkipCACheck);
        $script:session = New-PSSession -ConnectionURI:$ConnectionUri `
            -ConfigurationName:Microsoft.Exchange `
            -AllowRedirection `
            -Authentication:"Basic" `
            -SessionOption:$sessionOption `
            -Credential:$Credential `
            -ErrorAction:SilentlyContinue;
        
        if ($script:session -eq $null)
        {
            Write-Error ($LocalizedStrings.FailedToCreateRemoteSession -f $error[0].Exception.Message);
            Exit;
        }
        else
        {
            $result = Import-PSSession -Session $script:session `
                -Prefix "EXO" `
                -AllowClobber;

            if (-not $?)
            {
                Write-Error ($LocalizedStrings.FailedToImportRemoteSession -f $error[0].Exception.Message);
                Remove-PSSession $script:session;
                Exit;
            }
        }
    }
    finally
    {
        $WarningPreference = $oldWarningPreference;
        $VerbosePreference = $oldVerbosePreference;
    }

    WriteInfoMessage $LocalizedStrings.RemoteSessionCreatedSuccessfully;
}

# Invokes New-SyncMailPublicFolder to create a new MEPF object on AD
function NewMailEnabledPublicFolder()
{
    param ($localFolder)

    if ($localFolder.PrimarySmtpAddress.ToString() -eq "")
    {
        $errorMsg = ($LocalizedStrings.FailedToCreateMailPublicFolderEmptyPrimarySmtpAddress -f $localFolder.Guid);
        Write-Error $errorMsg;
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $errorMsg "";
        return;
    }

    # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddressesArray = $localFolder.EmailAddresses.ToStringArray() + ("x500:" + $localFolder.LegacyExchangeDN);
           
    $newParams = @{};
    AddNewOrSetCommonParameters $localFolder $emailAddressesArray $newParams;

    [string]$commandText = (FormatCommand $script:NewSyncMailPublicFolderCommand $newParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        $result = &$script:NewSyncMailPublicFolderCommand @newParams;
        WriteOperationSummary $localFolder $LocalizedStrings.CreateOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsCreated++;
        }
    }
    catch
    {
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $error[0].Exception.Message $commandText;
        Write-Error $_;
    }
}

# Invokes Remove-SyncMailPublicFolder to remove a MEPF from AD
function RemoveMailEnabledPublicFolder()
{
    param ($remoteFolder)    

    $removeParams = @{};
    $removeParams.Add("Identity", $remoteFolder.DistinguishedName);
    $removeParams.Add("Confirm", $false);
    $removeParams.Add("WarningAction", [System.Management.Automation.ActionPreference]::SilentlyContinue);
    $removeParams.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop);

    if ($WhatIf)
    {
        $removeParams.Add("WhatIf", $true);
    }
    
    [string]$commandText = (FormatCommand $script:RemoveSyncMailPublicFolderCommand $removeParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }
    
    try
    {
        &$script:RemoveSyncMailPublicFolderCommand @removeParams;
        WriteOperationSummary $remoteFolder $LocalizedStrings.RemoveOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsDeleted++;
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.RemoveOperationName $_.Exception.Message $commandText;
        Write-Error $_;
    }
}

# Invokes Set-MailPublicFolder to update the properties of an existing MEPF
function UpdateMailEnabledPublicFolder()
{
    param ($localFolder, $remoteFolder)

    $localEmailAddresses = $localFolder.EmailAddresses.ToStringArray();
    $localEmailAddresses += ("x500:" + $localFolder.LegacyExchangeDN); # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddresses = ConsolidateEmailAddresses $localEmailAddresses $remoteFolder.EmailAddresses $remoteFolder.LegacyExchangeDN;

    $setParams = @{};
    $setParams.Add("Identity", $remoteFolder.DistinguishedName);

    if ($script:mailEnabledSystemFolders.Contains($localFolder.Guid))
    {
        $setParams.Add("IgnoreMissingFolderLink", $true);
    }

    AddNewOrSetCommonParameters $localFolder $emailAddresses $setParams;

    [string]$commandText = (FormatCommand $script:SetMailPublicFolderCommand $setParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        &$script:SetMailPublicFolderCommand @setParams;
        WriteOperationSummary $remoteFolder $LocalizedStrings.UpdateOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsUpdated++;
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.UpdateOperationName $_.Exception.Message $commandText;
        Write-Error $_;
    }
}

# Adds the common set of parameters between New and Set cmdlets to the given dictionary
function AddNewOrSetCommonParameters()
{
    param ($localFolder, $emailAddresses, [System.Collections.IDictionary]$parameters)

    $windowsEmailAddress = $localFolder.WindowsEmailAddress.ToString();
    if ($windowsEmailAddress -eq "")
    {
        $windowsEmailAddress = $localFolder.PrimarySmtpAddress.ToString();      
    }

    $parameters.Add("Alias", $localFolder.Alias.Trim());
    $parameters.Add("DisplayName", $localFolder.DisplayName.Trim());
    $parameters.Add("EmailAddresses", $emailAddresses);
    $parameters.Add("ExternalEmailAddress", $localFolder.PrimarySmtpAddress.ToString());
    $parameters.Add("HiddenFromAddressListsEnabled", $localFolder.HiddenFromAddressListsEnabled);
    $parameters.Add("Name", $localFolder.Name.Trim());
    $parameters.Add("OnPremisesObjectId", $localFolder.Guid);
    $parameters.Add("WindowsEmailAddress", $windowsEmailAddress);
    $parameters.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop);

    if ($WhatIf)
    {
        $parameters.Add("WhatIf", $true);
    }
}

# Finds out the cloud-only email addresses and merges those with the values current persisted in the on-premises object
function ConsolidateEmailAddresses()
{
    param($localEmailAddresses, $remoteEmailAddresses, $remoteLegDN)

    # Check if the email address in the existing cloud object is present on-premises; if it is not, then the address was either:
    # 1. Deleted on-premises and must be removed from cloud
    # 2. or it is a cloud-authoritative address and should be kept
    $remoteAuthoritative = @();
    foreach ($remoteAddress in $remoteEmailAddresses)
    {
        if ($remoteAddress.StartsWith("SMTP:", [StringComparison]::InvariantCultureIgnoreCase))
        {
            $found = $false;
            $remoteAddressParts = $remoteAddress.Split($script:proxyAddressSeparators); # e.g. SMTP:alias@domain
            if ($remoteAddressParts.Length -ne 3)
            {
                continue; # Invalid SMTP proxy address (it will be removed)
            }

            foreach ($localAddress in $localEmailAddresses)
            {
                # note that the domain part of email addresses is case insensitive while the alias part is case sensitive
                $localAddressParts = $localAddress.Split($script:proxyAddressSeparators);
                if ($localAddressParts.Length -eq 3 -and
                    $remoteAddressParts[0].Equals($localAddressParts[0], [StringComparison]::InvariantCultureIgnoreCase) -and
                    $remoteAddressParts[1].Equals($localAddressParts[1], [StringComparison]::InvariantCulture) -and
                    $remoteAddressParts[2].Equals($localAddressParts[2], [StringComparison]::InvariantCultureIgnoreCase))
                {
                    $found = $true;
                    break;
                }
            }

            if (-not $found)
            {
                foreach ($domain in $script:authoritativeDomains)
                {
                    if ($remoteAddressParts[2] -eq $domain)
                    {
                        $found = $true;
                        break;
                    }
                }

                if (-not $found)
                {
                    # the address on the remote object is from a cloud authoritative domain and should not be removed
                    $remoteAuthoritative += $remoteAddress;
                }
            }
        }
        elseif ($remoteAddress.StartsWith("X500:", [StringComparison]::InvariantCultureIgnoreCase) -and
            $remoteAddress.Substring(5) -eq $remoteLegDN)
        {
            $remoteAuthoritative += $remoteAddress;
        }
    }

    return $localEmailAddresses + $remoteAuthoritative;
}

# Formats the command and its parameters to be printed on console or to file
function FormatCommand()
{
    param ([string]$command, [System.Collections.IDictionary]$parameters)

    $commandText = New-Object System.Text.StringBuilder;
    [void]$commandText.Append($command);
    foreach ($name in $parameters.Keys)
    {
        [void]$commandText.AppendFormat(" -{0}:",$name);

        $value = $parameters[$name];
        if ($value -isnot [Array])
        {
            [void]$commandText.AppendFormat("`"{0}`"", $value);
        }
        elseif ($value.Length -eq 0)
        {
            [void]$commandText.Append("@()");
        }
        else
        {
            [void]$commandText.Append("@(");
            foreach ($subValue in $value)
            {
                [void]$commandText.AppendFormat("`"{0}`",",$subValue);
            }
            
            [void]$commandText.Remove($commandText.Length - 1, 1);
            [void]$commandText.Append(")");
        }
    }

    return $commandText.ToString();
}

################ DECLARING GLOBAL VARIABLES ################
$script:session = $null;
$script:verbose = $VerbosePreference -eq [System.Management.Automation.ActionPreference]::Continue;

$script:csvSpecialChars = @("`r", "`n");
$script:csvEscapeChar = '"';
$script:csvFieldDelimiter = ',';

$script:ObjectsCreated = $script:ObjectsUpdated = $script:ObjectsDeleted = 0;
$script:NewSyncMailPublicFolderCommand = "New-EXOSyncMailPublicFolder";
$script:SetMailPublicFolderCommand = "Set-EXOMailPublicFolder";
$script:RemoveSyncMailPublicFolderCommand = "Remove-EXOSyncMailPublicFolder";
[char[]]$script:proxyAddressSeparators = ':','@';
$script:errorsEncountered = 0;
$script:authoritativeDomains = $null;
$script:mailEnabledSystemFolders = New-Object 'System.Collections.Generic.HashSet[Guid]'; 
$script:WellKnownSystemFolders = @(
    "\NON_IPM_SUBTREE\EFORMS REGISTRY",
    "\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK",
    "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
    "\NON_IPM_SUBTREE\schema-root",
    "\NON_IPM_SUBTREE\Events Root");

#load hashtable of localized string
Import-LocalizedData -BindingVariable LocalizedStrings -FileName SyncMailPublicFolders.strings.psd1

#minimum supported exchange version to run this script
$minSupportedVersion = 8
################ END OF DECLARATION #################

if (Test-Path $CsvSummaryFile)
{
    Remove-Item $CsvSummaryFile -Confirm:$Confirm -Force;
}

# Write the output CSV headers
$csvFile = New-Item -Path $CsvSummaryFile -ItemType File -Force -ErrorAction:Stop -Value ("#{0},{1},{2},{3},{4}`r`n" -f $LocalizedStrings.TimestampCsvHeader,
    $LocalizedStrings.IdentityCsvHeader,
    $LocalizedStrings.OperationCsvHeader,
    $LocalizedStrings.ResultCsvHeader,
    $LocalizedStrings.CommandCsvHeader);

$localServerVersion = (Get-ExchangeServer $env:COMPUTERNAME -ErrorAction:Stop).AdminDisplayVersion;
# This script can run from Exchange 2007 Management shell and above
if ($localServerVersion.Major -lt $minSupportedVersion)
{
    Write-Error ($LocalizedStrings.LocalServerVersionNotSupported -f $localServerVersion) -ErrorAction:Continue;
    Exit;
}

try
{
    InitializeExchangeOnlineRemoteSession;

    WriteInfoMessage $LocalizedStrings.LocalMailPublicFolderEnumerationStart;

    # During finalization, Public Folders deployment is locked for migration, which means the script cannot invoke
    # Get-PublicFolder as that operation would fail. In that case, the script cannot determine which mail public folder
    # objects are linked to system folders under the NON_IPM_SUBTREE.
    $lockedForMigration = (Get-OrganizationConfig).PublicFoldersLockedForMigration;
    $allSystemFoldersInAD = @();
    if (-not $lockedForMigration)
    {
        # See https://technet.microsoft.com/en-us/library/bb397221(v=exchg.141).aspx#Trees
        # Certain WellKnownFolders in pre-E15 are created with prefix such as OWAScratchPad, StoreEvents.
        # For instance, StoreEvents folders have the following pattern: "\NON_IPM_SUBTREE\StoreEvents{46F83CF7-2A81-42AC-A0C6-68C7AA49FF18}\internal1"
        $storeEventAndOwaScratchPadFolders = @(Get-PublicFolder \NON_IPM_SUBTREE -GetChildren -ResultSize:Unlimited | ?{$_.Name -like "StoreEvents*" -or $_.Name -like "OWAScratchPad*"});
        $allSystemFolderParents = $storeEventAndOwaScratchPadFolders + @($script:WellKnownSystemFolders | Get-PublicFolder -ErrorAction:SilentlyContinue);
        $allSystemFoldersInAD = @($allSystemFolderParents | Get-PublicFolder -Recurse -ResultSize:Unlimited | Get-MailPublicFolder -ErrorAction:SilentlyContinue);

        foreach ($systemFolder in $allSystemFoldersInAD)
        {
            [void]$script:mailEnabledSystemFolders.Add($systemFolder.Guid);
        }
    }
    else
    {
        WriteWarningMessage $LocalizedStrings.UnableToDetectSystemMailPublicFolders;
    }

    if ($script:verbose)
    {
        WriteVerboseMessage ($LocalizedStrings.SystemFoldersSkipped -f $script:mailEnabledSystemFolders.Count);
        $allSystemFoldersInAD | Sort Alias | ft -a | Out-String | Write-Host -ForegroundColor Green -BackgroundColor Black;
    }

    $localFolders = @(Get-MailPublicFolder -ResultSize:Unlimited -IgnoreDefaultScope | Sort Guid);
    WriteInfoMessage ($LocalizedStrings.LocalMailPublicFolderEnumerationCompleted -f $localFolders.Length);

    if ($localFolders.Length -eq 0 -and $Force -eq $false)
    {
        WriteWarningMessage $LocalizedStrings.ForceParameterRequired;
        Exit;
    }

    WriteInfoMessage $LocalizedStrings.RemoteMailPublicFolderEnumerationStart;
    $remoteFolders = @(Get-EXOMailPublicFolder -ResultSize:Unlimited | Sort OnPremisesObjectId);
    WriteInfoMessage ($LocalizedStrings.RemoteMailPublicFolderEnumerationCompleted -f $remoteFolders.Length);

    $missingOnPremisesGuid = @();
    $pendingRemoves = @();
    $pendingUpdates = @{};
    $pendingAdds = @{};

    $localIndex = 0;
    $remoteIndex = 0;
    while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $remoteFolders.Length)
    {
        $local = $localFolders[$localIndex];
        $remote = $remoteFolders[$remoteIndex];

        if ($remote.OnPremisesObjectId -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote;
            $remoteIndex++;
        }
        elseif ($local.Guid.ToString() -eq $remote.OnPremisesObjectId)
        {
            $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local=$local; Remote=$remote }));
            $localIndex++;
            $remoteIndex++;
        }
        elseif ($local.Guid.ToString() -lt $remote.OnPremisesObjectId)
        {
            if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
            {
                $pendingAdds.Add($local.Guid, $local);
            }

            $localIndex++;
        }
        else
        {
            $pendingRemoves += $remote;
            $remoteIndex++;
        }
    }

    # Remaining folders on $localFolders collection must be added to Exchange Online
    while ($localIndex -lt $localFolders.Length)
    {
        $local = $localFolders[$localIndex];

        if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
        {
            $pendingAdds.Add($local.Guid, $local);
        }

        $localIndex++;
    }

    # Remaining folders on $remoteFolders collection must be removed from Exchange Online
    while ($remoteIndex -lt $remoteFolders.Length)
    {
        $remote = $remoteFolders[$remoteIndex];
        if ($remote.OnPremisesObjectId  -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote;
        }
        else
        {
            $pendingRemoves += $remote;
        }
        
        $remoteIndex++;
    }

    if ($missingOnPremisesGuid.Length -gt 0)
    {
        # Process remote objects missing the OnPremisesObjectId using the PrimarySmtpAddress as a key instead.
        $missingOnPremisesGuid = @($missingOnPremisesGuid | Sort PrimarySmtpAddress);
        $localFolders = @($localFolders | Sort PrimarySmtpAddress);

        $localIndex = 0;
        $remoteIndex = 0;
        while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $local = $localFolders[$localIndex];
            $remote = $missingOnPremisesGuid[$remoteIndex];

            if ($local.PrimarySmtpAddress.ToString() -eq $remote.PrimarySmtpAddress.ToString())
            {
                # Make sure the PrimarySmtpAddress has no duplicate on-premises; otherwise, skip updating all objects with duplicate address
                $j = $localIndex + 1;
                while ($j -lt $localFolders.Length)
                {
                    $next = $localFolders[$j];
                    if ($local.PrimarySmtpAddress.ToString() -ne $next.PrimarySmtpAddress.ToString())
                    {
                        break;
                    }

                    WriteErrorSummary $next $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByAnotherFolder -f $local.PrimarySmtpAddress,$local.Guid) "";

                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($next.Guid);
                    $pendingUpdates.Remove($next.Guid);
                    $j++;
                }

                $duplicatesFound = $j - $localIndex - 1;
                if ($duplicatesFound -gt 0)
                {
                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($local.Guid);
                    $pendingUpdates.Remove($local.Guid);
                    $localIndex += $duplicatesFound + 1;

                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByOtherFolders -f $local.PrimarySmtpAddress,$duplicatesFound) "";
                    WriteWarningMessage ($LocalizedStrings.SkippingFoldersWithDuplicateAddress -f ($duplicatesFound + 1),$local.PrimarySmtpAddress);
                }
                elseif ($pendingUpdates.Contains($local.Guid))
                {
                    # If we get here, it means two different remote objects match the same local object (one by OnPremisesObjectId and another by PrimarySmtpAddress).
                    # Since that is an ambiguous resolution, let's skip updating the remote objects.
                    $ambiguousRemoteObj = $pendingUpdates[$local.Guid].Remote;
                    $pendingUpdates.Remove($local.Guid);

                    $errorMessage = ($LocalizedStrings.AmbiguousLocalMailPublicFolderResolution -f $local.Guid,$ambiguousRemoteObj.Guid,$remote.Guid);
                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName $errorMessage "";
                    WriteWarningMessage $errorMessage;
                }
                else
                {
                    # Since there was no match originally using OnPremisesObjectId, the local object was treated as an add to Exchange Online.
                    # In this way, since we now found a remote object (by PrimarySmtpAddress) to update, we must first remove the local object from the add list.
                    $pendingAdds.Remove($local.Guid);
                    $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local=$local; Remote=$remote }));
                }

                $localIndex++;
                $remoteIndex++;
            }
            elseif ($local.PrimarySmtpAddress.ToString() -gt $remote.PrimarySmtpAddress.ToString())
            {
                # There are no local objects using the remote object's PrimarySmtpAddress
                $pendingRemoves += $remote;
                $remoteIndex++;
            }
            else
            {
                $localIndex++;
            }
        }

        # All objects remaining on the $missingOnPremisesGuid list no longer exist on-premises
        while ($remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $pendingRemoves += $missingOnPremisesGuid[$remoteIndex];
            $remoteIndex++;
        }
    }

    $script:totalItems = $pendingRemoves.Length + $pendingUpdates.Count + $pendingAdds.Count;

    # At this point, we know all changes that need to be synced to Exchange Online. Let's prompt the admin for confirmation before proceeding.
    if ($Confirm -eq $true -and $script:totalItems -gt 0)
    {
        $title = $LocalizedStrings.ConfirmationTitle;
        $message = ($LocalizedStrings.ConfirmationQuestion -f $pendingAdds.Count,$pendingUpdates.Count,$pendingRemoves.Length);
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationYesOption, `
            $LocalizedStrings.ConfirmationYesOptionHelp;

        $no = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationNoOption, `
            $LocalizedStrings.ConfirmationNoOptionHelp;

        [System.Management.Automation.Host.ChoiceDescription[]]$options = $no,$yes;
        $confirmation = $host.ui.PromptForChoice($title, $message, $options, 0);
        if ($confirmation -eq 0)
        {
            Exit;
        }
    }

    # Find out the authoritative AcceptedDomains on-premises so that we don't accidently remove cloud-only email addresses during updates
    $script:authoritativeDomains = @(Get-AcceptedDomain | ?{$_.DomainType -eq "Authoritative" } | foreach {$_.DomainName.ToString()});
    
    # Finally, let's perfom the actual operations against Exchange Online
    $script:itemsProcessed = 0;
    for ($i = 0; $i -lt $pendingRemoves.Length; $i++)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusRemoving $i $pendingRemoves.Length;
        RemoveMailEnabledPublicFolder $pendingRemoves[$i];
    }

    $script:itemsProcessed += $pendingRemoves.Length;
    $updatesProcessed = 0;
    foreach ($folderPair in $pendingUpdates.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusUpdating $updatesProcessed $pendingUpdates.Count;
        UpdateMailEnabledPublicFolder $folderPair.Local $folderPair.Remote;
        $updatesProcessed++;
    }

    $script:itemsProcessed += $pendingUpdates.Count;
    $addsProcessed = 0;
    foreach ($localFolder in $pendingAdds.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusCreating $addsProcessed $pendingAdds.Count;
        NewMailEnabledPublicFolder $localFolder;
        $addsProcessed++;
    }

    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity -Status ($LocalizedStrings.ProgressBarStatusCreating -f $pendingAdds.Count,$pendingAdds.Count) -Completed;
    WriteInfoMessage ($LocalizedStrings.SyncMailPublicFolderObjectsComplete -f $script:ObjectsCreated,$script:ObjectsUpdated,$script:ObjectsDeleted);

    if ($script:errorsEncountered -gt 0)
    {
        WriteWarningMessage ($LocalizedStrings.ErrorsFoundDuringImport -f $script:errorsEncountered,(Get-Item $CsvSummaryFile).FullName);
    }
}
finally
{
    if ($script:session -ne $null)
    {
        Remove-PSSession $script:session;
    }
}
# SIG # Begin signature block
# MIIdtAYJKoZIhvcNAQcCoIIdpTCCHaECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHZbIYGfoB0qYph1U53JyqTk9
# B9agghhkMIIEwzCCA6ugAwIBAgITMwAAAJgEWMt/IwmwngAAAAAAmDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI3
# WhcNMTcwNjMwMTkyMTI3WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjdBRkEtRTQxQy1FMTQyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA1jclqAQB7jVZ
# CvOuH5jFixrRTGFtwMHws1sEZaA3ciobVIdWIejc5fBu3XdwRLfxjsmyou3JeTaa
# 8lqA929q2AyZ9A3ZBfxf8VqTxbu06wBj4o4g5YCsz0C/81N2ESsQZbjDxbW5sKzD
# hhT0nTzr82aepe1drjT5dvyU/AvEkCzaEDU0dZTq2Bm6NIif11GzA+OkD0bdZG+u
# 4EJRylQ4ijStGgXUpAapb0y2RtlAYvZSpLYzeFFcA/yRXacCnoD++h9r66he/Scv
# Gfd/J/5hPRCtgsbNr3vFJzBWgV9zVqmWOvZBPGpLhCLglTh0stPa/ZxZjTS/nKJL
# a7MZId131QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFPPCI5/SvSWNvaj1nBvoSHO7
# 6ZPBMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAD+xPVIhFl30XEe39rlgUqCCr2fXR9o0aL0Oioap6LAUMXLK
# 4B+/L2c+BgV32joU6vMChTaA+7XEw7pXCRN+uD8ul4ifHrdAOEEqOTBD7N5203u2
# LN667/WY71purP2ezNB1y+YAgjawEt6VjjQcSGZ9bTPRtS2JPS5BS868paym355u
# 16HMxwmhlv1klX6nfVOs1DYK5cZUrPAblCZEWzGab8j9d2ZIGLQmTEmStdslOq79
# vujEI0nqDnJBusUGi28Kh3Hz1QIHM5UZg/F5sWgt0EobFGHmk4KH2vreGZArtCIB
# amDc5cIJ48na9GfA2jqJLWsbvNcwC486g5cauwkwggYHMIID76ADAgECAgphFmg0
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUHTHPErF64W0nw+pkJsWDIg9obn0wbgYKKwYB
# BAGCNwIBDDFgMF6gNoA0AFMAeQBuAGMALQBNAGEAaQBsAFAAdQBiAGwAaQBjAEYA
# bwBsAGQAZQByAHMALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20v
# ZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAHv4sjSDzgC+2FGAZk6TmVDbbebI
# 4wqa8xvOTaMPQC3FTvJ4V/wAsOw7x5BTNe9cCvuO5+KWl7+U3duw2Ee9t6mhSBz9
# XvuzAULLP4OifMsb6jli/BoCz0v8ZhpvRopFA+nbO7HrhElz2mUfn1IgW0+XjNwi
# E+V7jV0eeOBHQnJRs3sSHNihqxkRRTcg5xPakRNsLxuDOa/OeC8h1G/spFauhPgE
# qf7i/WXhhhi+Ky10jTdD5x7aBG6dlYLIySEXEhQml7mS/PI4uM9UDhuv0bBkl4cb
# OgJ2u+2RzvO1hEhG3v/w1RTIwm+vH89aRiu3pR+gvyhaBZPH8Z5Ur7lTZEuhggIo
# MIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBAhMzAAAAmARYy38jCbCeAAAAAACYMAkGBSsOAwIaBQCgXTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ0MzFa
# MCMGCSqGSIb3DQEJBDEWBBRE3kjbJ+UOg9WwJ5c0VGW0uorBZDANBgkqhkiG9w0B
# AQUFAASCAQBpmtTs/rpOugNM8Y0JNJOXnTc+Wfa7jNV/Rl1kaMAdwEZtb9kA+Kcx
# 548r07qbjAZGIF4CyQ7gE/aSZW5HomphaFFM3enZNSQx0y2+k6W9aMHfqKmnUerB
# 8Y7E/izsIvK/EvKPixngXigbMZYEbx42WAoZeFMMRfVBqHg2RBrL579JdovcDTYS
# bQ8bAvV2656KzHMz3s5EWKqv4vEvAepl/mCPD7S6M4kYu3DVXaQciNLDSOXsxyKp
# jxcLxUP4sem+/GnExgPgggCQFODrBuCpgEqpPzDfIZmLpk/H20RQ0P1qQnpU0QEy
# xgVU65IlffGWH23uAY1xRJ7vsGmcRAmt
# SIG # End signature block
