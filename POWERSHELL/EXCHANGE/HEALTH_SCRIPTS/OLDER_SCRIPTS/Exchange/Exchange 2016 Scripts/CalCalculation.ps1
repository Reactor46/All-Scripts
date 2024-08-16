########################
##  Input parameters  ##
########################
[CmdletBinding(SupportsShouldProcess = $false, ConfirmImpact = "None", DefaultParameterSetName="Default")]
param
(
    [Parameter(Position = 0, Mandatory = $false, ValueFromPipeline = $true)]
    [ValidateSet(15)]
    [int]
    $VersionMajor = 15,

    [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline = $true)]
    [ValidateSet("Summary", "Standard", "Enterprise")]
    [string]
    $AccessLicenseType = "Summary",

    [Parameter(Position = 2, Mandatory = $false, ValueFromPipeline = $true)]
    [ValidateSet("All", "Journaling", "ActiveSync", "UM","ManagedFolder", "RetentionPolicy", "PersonalArchive", "LegalHold", "DLP")]
    [string[]]
    $DebugCategory = "All",

    [Parameter(Position = 3, Mandatory = $false, ValueFromPipeline = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DebugMailbox = $null
)

############################
## Script level variables ##
############################
$Script:TotalMailboxes = 0
$Script:TotalEnterpriseCALs = 0
$Script:OrgWideJournalingEnabled = $False
$Script:AllMailboxIDs = @{}
$Script:AllVersionMailboxIDs = @{}
$Script:EnterpriseCALMailboxIDs = @{}
$Script:JournalingUserCount = 0
$Script:JournalingMailboxIDs = @{}
$Script:JournalingDGMailboxMemberIDs = @{}
$Script:TotalStandardCALs = 0
$Script:VisitedGroups = @{}
$Script:DGStack = new-object System.Collections.Stack
$Script:UserMailboxFilter = "(RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox') -or (RecipientTypeDetails -eq 'LinkedMailbox')"

$Script:ManagedFolderMailboxPolicyWithCustomedFolder = @{}
$Script:RetentionPolicyWithPersonalTag = @{}
$Script:RetentionPolicyWithPersonalTagNonArchive = @{}
$Script:ActiveSyncMailboxPolicyWithECALFeature = @{}
$Script:DebugMailboxGuid = $null
$Script:bacupErrorActionPreference = "Continue"

#######################
## Exception handler ##
#######################
# Error handling in this script:
# 1. If any error happend, let's just stop the rest since all the data will be marked as failed if one error happened. So $ErrorActionPreference is set to 'Stop'
# 2. If -ErrorAction SilentlyContinue is specified, which means any errors during that cmdlet ared not considered as error, but $error still records them, so make sure
#    $error.Clear() is called after cmdlet with -ErrorAction SilentlyContinue

# Trap block
trap
{
    Restore-EnvironmentVariable
    exit
}

####################
## Log functions  ##
####################
# Log when verbose is on
function Log-Info([string] $info)
{
    Write-Verbose (FormatEntry $info)
}

# Log when debug is on
function Log-Debug([string] $info)
{
    Write-Debug (FormatEntry $info)
}

# Log enter when verbose is on
function Log-Enter([string] $name)
{
    Log-Info "$name : Beginning processing"
}

# Log exit when verbose in on
function Log-Exit([string] $name)
{
    Log-Info "$name : Ending processing"
}

# Format log entry
function FormatEntry([string] $info)
{
    return "[{0} UTC] {1}" -F $(get-date).ToUniversalTime().ToString("HH:mm:ss.4d"), $info
}

######################
## Output functions ##
######################
# Output objects to caller
function Output-Report
{
    if ($Script:AccessLicenseType -eq "Summary")
    {
        write-output $Script:TotalMailboxes
        write-output $Script:TotalStandardCALs
        write-output $Script:TotalEnterpriseCALs
        Write-output $Script:JournalingUserCount
    }
    elseif ($Script:AccessLicenseType -eq "Standard")
    {
        $Script:AllMailboxIDs.Values | foreach {
            Write-output $_
        }
    }
    elseif ($Script:AccessLicenseType -eq "Enterprise")
    {
        $Script:EnterpriseCALMailboxIDs.Values | foreach {
            Write-output $_
        }
    }
    else
    {
        Throw "AccessLicenseType: $Script:AccessLicenseType is not supported."
    }
}

####################################
## Helper functions for Debugging ##
####################################
# Function that checks if we should process for this category
function Process-Category([string] $category)
{
    if (($Script:DebugCategory -contains "All") -or ($Script:DebugCategory -contains $category))
    {
        Log-Info "Will process category $category"

        return $true
    }
    else
    {
        Log-Info "Will Ignore category $category"

        return $false
    }
}

# Set debug mailbox guid by its address
function Set-DebugMailboxGuid
{
    if (($Script:DebugMailbox -eq $null) -or ($Script:DebugMailbox -eq ""))
    {
        $Script:DebugMailboxGuid = $null
    }
    elseif ($Script:DebugMailbox -ieq "All")
    {
        $Script:DebugMailboxGuid = "All"

        Log-Debug "DebugMailbox is set for all mailboxes. DebugMailbox info will only show with -Verbose on."
    }
    else
    {
        $mailbox = Get-Mailbox -Identity $Script:DebugMailbox
        if ($mailbox -eq $null)
        {
            Throw "The mailbox $Script:DebugMailbox does not exist."
        }

        $Script:DebugMailboxGuid = $mailbox.Guid

        Log-Debug "The Guid of DebugMailbox ($Script:DebugMailbox) is $Script:DebugMailboxGuid"
    }
}

# Determine if we should set and log this mailbox info or not.
function Process-Mailbox([string] $guid, [string] $varName)
{
    if ($Script:DebugMailboxGuid -eq $null)
    {
        # Not debugging. Don't log.
        return $true
    }
    elseif ($Script:DebugMailboxGuid -ieq "All")
    {
        # Debug all. Log all using verbose.
        Log-Info "DebugMailbox All (current: $guid) is in $varName"
        return $true
    }
    else
    {
        if ($guid -ieq $Script:DebugMailboxGuid)
        {
            # Debug one. Log when matched using debug
            Log-Debug "DebugMailbox ($guid) is in $varName"
            return $true
        }
        else
        {
            # Debug one. No match, so ignore it.
            Log-Info "DebugMailbox ($Script:DebugMailboxGuid) $guid is Ignored for $varName"
            return $false
        }
    }
}

######################
## Helper functions ##
######################
# Function that merges two hashtables
function Merge-Hashtables
{
    $Table1 = $args[0]
    $Table2 = $args[1]
    $Result = @{}
    
    if ($null -ne $Table1)
    {
        $Result += $Table1
    }

    if ($null -ne $Table2)
    {
        foreach ($entry in $Table2.GetEnumerator())
        {
            $Result[$entry.Key] = $entry.Value
        }
    }

    $Result
}

# Function that returns the value for output
function Get-MailboxOutputValue($mailbox)
{
    return $mailbox | Select PrimarySmtpAddress
}

# Help function for function Get-JournalingGroupMailboxMember to traverse members of a DG/DDG/group 
function Traverse-GroupMember
{
    $GroupMember = $args[0]
    
    if( $GroupMember -eq $null )
    {
        return
    }

    # Note!!! 
    # Only user, shared and linked mailboxes are counted. 
    # Resource mailboxes and legacy mailboxes are NOT counted.
    if ( ($GroupMember.RecipientTypeDetails -eq 'UserMailbox') -or
          ($GroupMember.RecipientTypeDetails -eq 'SharedMailbox') -or
          ($GroupMember.RecipientTypeDetails -eq 'LinkedMailbox') ) {
        # Journal one mailbox
        if (Process-Mailbox $GroupMember.Guid "JournalingMailboxIDs")
        {
            $Script:JournalingMailboxIDs[$GroupMember.Guid] = $null
        }
    } elseif ( ($GroupMember.RecipientType -eq "Group") -or ($GroupMember.RecipientType -like "Dynamic*Group") -or ($GroupMember.RecipientType -like "Mail*Group") ) {
        Log-Info "Push this DG/DDG/group into the stack. ($GroupMember.Guid)"
        $Script:DGStack.Push(@($GroupMember.Guid, $GroupMember.RecipientType))
    }
}

# Function that returns all mailbox members including duplicates recursively from a DG/DDG
function Get-JournalingGroupMailboxMember
{
    # Skip this DG/DDG if it was already enumerated.
    if ( $Script:VisitedGroups.ContainsKey($args[0]) ) {
        return
    }
    
    $Script:DGStack.Push(@($args[0],$args[1]))
    while ( $Script:DGStack.Count -ne 0 ) {
        $StackElement = $DGStack.Pop()
        
        $GroupGuid = $StackElement[0]
        $GroupRecipientType = $StackElement[1]

        if ( $Script:VisitedGroups.ContainsKey($GroupGuid) ) {
            # Skip this this DG/DDG if it was already enumerated.
            continue
        }
        
        Log-Info "Check the members of the current DG/DDG/group in the stack. ($GroupGuid)"
        if ( ($GroupRecipientType -like "Mail*Group") -or ($GroupRecipientType -eq "Group" ) ) {
            $varGroup = Get-Group $GroupGuid.ToString() -ErrorAction SilentlyContinue
            $error.Clear()
            if ( $varGroup -eq $Null )
            {
                return
            }
            
            $varGroup.members | foreach {    
                # Count users and groups which could be mailboxes.
                $varGroupMember = Get-User $_ -ErrorAction SilentlyContinue 
                if ( $varGroupMember -eq $Null ) {
                    $varGroupMember = Get-Group $_ -ErrorAction SilentlyContinue                  
                }
                $error.Clear()

                if ( $varGroupMember -ne $Null ) {
                    Traverse-GroupMember $varGroupMember
                }
            }
        } else {
            Log-Info "The current stack element is a DDG. ($GroupGuid)"
            $varGroup = Get-DynamicDistributionGroup $GroupGuid.ToString() -ErrorAction SilentlyContinue
            $error.Clear()

            if ( $varGroup -eq $Null )
            {
                return
            }

            Get-Recipient -RecipientPreviewFilter $varGroup.LdapRecipientFilter -OrganizationalUnit $varGroup.RecipientContainer -ResultSize 'Unlimited' -PropertySet 'Minimum' | foreach {
                Traverse-GroupMember $_
            }
        } 

        # Mark this DG/DDG as visited as it's enumerated.
        $Script:VisitedGroups[$GroupGuid] = $null
    }    
}

# Backup and set powershell environment variable
function BackupSet-EnvironmentVariable
{
    $Script:bacupErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
}

# Restore powershell environment variable
function Restore-EnvironmentVariable
{
    $ErrorActionPreference = $Script:bacupErrorActionPreference
}

########################################
## Calculate Standard CAL functions   ##
########################################
#
# Calc total # of mailboxes
#
function Calc-AllMaiboxes
{
    Log-Enter $MyInvocation.MyCommand

    Log-Info "Calc-AllMaiboxes: Only user, shared and linked mailboxes are counted. Resource mailboxes, legacy mailboxes and team mailboxes are NOT counted."
    Log-Info "UserMailboxFilter : $UserMailboxFilter"
    Get-Recipient -ResultSize 'Unlimited' -Filter $UserMailboxFilter -PropertySet 'Minimum' | foreach {
        $Mailbox = $_
        if ($Mailbox.ExchangeVersion.ExchangeBuild.Major -eq $Script:VersionMajor) {
            if (Process-Mailbox $Mailbox.Guid "AllMailboxIDs")
            {
                $Script:AllMailboxIDs[$Mailbox.Guid] = Get-MailboxOutputValue $Mailbox
                $Script:TotalMailboxes++
            }
        }

        if (Process-Mailbox $Mailbox.Guid "AllVersionMailboxIDs")
        {
            $Script:AllVersionMailboxIDs[$Mailbox.Guid] = Get-MailboxOutputValue $Mailbox
        }
    }
    Log-Debug "AllVersionMailboxIDs count is: $($AllVersionMailboxIDs.Count)"
    Log-Debug "Calc TotalMailboxs is: $Script:TotalMailboxes"
    Log-Exit $MyInvocation.MyCommand
}

####################################################
## Cache functions for Enterprise CAL Calculation ##
####################################################
#
# Cache for ManagedFolder
#
function Cache-ForManagedFolder
{
    Log-Enter $MyInvocation.MyCommand

	try
	{
		if (Process-Category "ManagedFolder")
		{
			# Setup cache for MRM to reduce task call times.
			Get-ManagedFolderMailboxPolicy | foreach {
				foreach ($FolderId in $_.ManagedFolderLinks)
				{
					$ManagedFolder = Get-ManagedFolder $FolderId
					if ($ManagedFolder.FolderType -eq "ManagedCustomFolder")
					{
						if (Process-Mailbox $_.Guid "ManagedFolderMailboxPolicyWithCustomedFolder")
						{
							$Script:ManagedFolderMailboxPolicyWithCustomedFolder[$_.Guid] = $null
							break
						}
					}
				}
			}

			Log-Debug "Cached ManagedFolderMailboxPolicyWithCustomedFolder count is: $($Script:ManagedFolderMailboxPolicyWithCustomedFolder.Count)"
		}
	}
	catch
	{
        Log-Debug "Cached ManagedFolderMailboxPolicyWithCustomedFolder failed"
	}

    Log-Exit $MyInvocation.MyCommand
}

#
# Cache for RetentionPolicy
#
function Cache-ForRetentionPolicy
{
    Log-Enter $MyInvocation.MyCommand

    if (Process-Category "RetentionPolicy")
    {
        $retentionPolicies = Get-RetentionPolicy
        $retentionPolicies | foreach {
            foreach ($PolicyTagID in $_.RetentionPolicyTagLinks) {
                $RetentionPolicyTag = Get-RetentionPolicyTag $PolicyTagID
                if ($RetentionPolicyTag.Type -eq "Personal")
                {
                    if (Process-Mailbox $_.Guid "RetentionPolicyWithPersonalTag")
                    {
                        $Script:RetentionPolicyWithPersonalTag[$_.Guid] = $null
                    }

                    if ($RetentionPolicyTag.RetentionAction -ne "MoveToArchive")
                    {
                        if (Process-Mailbox $_.Guid "RetentionPolicyWithPersonalTagNonArchive")
                        {
                            $Script:RetentionPolicyWithPersonalTagNonArchive[$_.Guid] = $null
                        }
                        break;
                    }
                }
            }
        }

        Log-Debug "Cached RetentionPolicyWithPersonalTag count is: $($Script:RetentionPolicyWithPersonalTag.Count)"
        Log-Debug "Cached RetentionPolicyWithPersonalTagNonArchive count is: $($Script:ActiveSyncMailboxPolicyWithECALFeature.Count)"
    }

    Log-Exit $MyInvocation.MyCommand
}

#
# Cache for ActiveSync
#
function Cache-ForActiveSync
{
    Log-Enter $MyInvocation.MyCommand

    if (Process-Category "ActiveSync")
    {
        # Setup cache to reduce Get-ActiveSyncMailboxPolicy.
        Get-ActiveSyncMailboxPolicy | foreach {
            $ASPolicy = $_
            if (($ASPolicy.AllowDesktopSync -eq $False) -or 
                    ($ASPolicy.AllowStorageCard -eq $False) -or
                    ($ASPolicy.AllowCamera -eq $False) -or
                    ($ASPolicy.AllowTextMessaging -eq $False) -or
                    ($ASPolicy.AllowWiFi -eq $False) -or
                    ($ASPolicy.AllowBluetooth -ne "Allow") -or
                    ($ASPolicy.AllowIrDA -eq $False) -or
                    ($ASPolicy.AllowInternetSharing -eq $False) -or
                    ($ASPolicy.AllowRemoteDesktop -eq $False) -or
                    ($ASPolicy.AllowPOPIMAPEmail -eq $False) -or
                    ($ASPolicy.AllowConsumerEmail -eq $False) -or
                    ($ASPolicy.AllowBrowser -eq $False) -or
                    ($ASPolicy.AllowUnsignedApplications -eq $False) -or
                    ($ASPolicy.AllowUnsignedInstallationPackages -eq $False) -or
                    ($ASPolicy.ApprovedApplicationList -ne $null) -or
                    ($ASPolicy.UnapprovedInROMApplicationList -ne $null))
                    {
                        if (Process-Mailbox $ASPolicy.Guid "ActiveSyncMailboxPolicyWithECALFeature")
                        {
                            $Script:ActiveSyncMailboxPolicyWithECALFeature[$ASPolicy.Guid] = $null
                        }
                    }
        }

        Log-Debug "Cached ActiveSyncMailboxPolicyWithECALFeature count is: $($Script:ActiveSyncMailboxPolicyWithECALFeature.Count)"
    }

    Log-Exit $MyInvocation.MyCommand
}

###############################################
## Calculate Enterprise CAL (ECAL) functions ##
###############################################
#
# Per-org Enterprise CALs
#
function Calc-ECALForOrg
{
    Log-Enter $MyInvocation.MyCommand

    $ret = $false

    # Consider this belongs to DLP
    if (Process-Category "DLP")
    {
        # If any RMS transport rule is defined, all mailboxes in the org are counted as Enterprise CALs.
        foreach($rule in Get-TransportRule)
        {
            if ($rule.ApplyRightsProtectionTemplate -ne $null) {
                $Script:TotalEnterpriseCALs = $Script:TotalMailboxes
                $Script:EnterpriseCALMailboxIDs = $Script:AllMailboxIDs

                # All mailboxes are counted as Enterprise CALs
                $ret = $true
                break;
            }
        }

        Log-Debug "Calc DLP count is: $Script:TotalEnterpriseCALs"
    }

    Log-Exit $MyInvocation.MyCommand

    return $ret
}

#
# Calculate Enterprise CAL users for UM, MRM Managed Custom Folder, and advanced ActiveSync policy and Legal Hold
# AVOID call task directly in the loop which can task consuming with large organization.
#
function Calc-ECALForMultiple
{
    Log-Enter $MyInvocation.MyCommand

    if ((Process-Category "UM") -or
        (Process-Category "PersonalArchive") -or
        (Process-Category "RetentionPolicy") -or
        (Process-Category "ManagedFolder") -or
        (Process-Category "LegalHold"))
    {
        $UMCount = 0
        $PersonalArchiveCount = 0
        $RetentionPolicyCount = 0
        $ManagedFolderCount = 0
        $LegalHoldCount = 0

        # AVOID call task directly in the loop which can task consuming with large organization.
        Get-Recipient -ResultSize 'Unlimited' -Filter $UserMailboxFilter -PropertySet 'ConsoleLargeSet' | foreach {  
            $Mailbox = $_
            if ($Mailbox.ExchangeVersion.ExchangeBuild.Major -eq $Script:VersionMajor)
            {
                # UM usage classifies the user as an Enterprise CAL   
                if ((Process-Category "UM") -and $Mailbox.UMEnabled)
                {
                    if (Process-Mailbox $Mailbox.Guid "UMEnabled")
                    {
                        $UMCount++
                        $Script:EnterpriseCALMailboxIDs[$Mailbox.Guid] = Get-MailboxOutputValue $Mailbox
                        return
                    }
                }

                # LOCAL Archive Mailbox classifies the user as an Enterprise CAL
                if ((Process-Category "PersonalArchive") -and ($Mailbox.ArchiveState -eq "Local"))
                {
                    if (Process-Mailbox $Mailbox.Guid "PersonalArchive")
                    {
                        $PersonalArchiveCount++
                        $Script:EnterpriseCALMailboxIDs[$Mailbox.Guid] = Get-MailboxOutputValue $Mailbox
                        return
                    }
                }
        
                # Retention Policy classifies the user as an Enterprise CAL
                if ((Process-Category "RetentionPolicy") -and
                    ($Mailbox.RetentionPolicy -ne $null) -and
                    $Script:RetentionPolicyWithPersonalTag.Contains($Mailbox.RetentionPolicy.ObjectGuid))
                {
                    # For online archive, we will not consider it as ECAL if it's caused by MoveToAchiveTag
                    if (($Mailbox.ArchiveState -eq "HostedProvisioned") -or ($Mailbox.ArchiveState -eq "HostedPending"))
                    {
                        if ($Script:RetentionPolicyWithPersonalTagNonArchive.Contains($Mailbox.RetentionPolicy.ObjectGuid))
                        {
                            if (Process-Mailbox $Mailbox.Guid "RetentionPolicy")
                            {
                                $RetentionPolicyCount++
                                $Script:EnterpriseCALMailboxIDs[$Mailbox.Guid] = Get-MailboxOutputValue $Mailbox
                                return
                            }
                        }
                    }
                    else
                    {
                        if (Process-Mailbox $Mailbox.Guid "RetentionPolicy")
                        {
                            $RetentionPolicyCount++
                            $Script:EnterpriseCALMailboxIDs[$Mailbox.Guid] = Get-MailboxOutputValue $Mailbox
                            return
                        }
                    }
                }

                # MRM Managed Custom Folder usage classifies the user as an Enterprise CAL
                if ((Process-Category "ManagedFolder") -and
                    ($Mailbox.ManagedFolderMailboxPolicy -ne $null) -and           
                    ($Script:ManagedFolderMailboxPolicyWithCustomedFolder.Contains($Mailbox.ManagedFolderMailboxPolicy.ObjectGuid)))
                {
                    if (Process-Mailbox $Mailbox.Guid "ManagedFolderMailboxPolicy")
                    {
                        $ManagedFolderCount++
                        $Script:EnterpriseCALMailboxIDs[$Mailbox.Guid] = Get-MailboxOutputValue $Mailbox
                        return						
                    }
                }

                # LitigationHoldEnabled Mailbox classifies the user as an Enterprise CAL
                if ((Process-Category "LegalHold") -and $Mailbox.LitigationHoldEnabled)
                {
                    if (Process-Mailbox $Mailbox.Guid "LitigationHoldEnabled")
                    {
                        $LegalHoldCount++
                        $Script:EnterpriseCALMailboxIDs[$Mailbox.Guid] = Get-MailboxOutputValue $Mailbox
                        return
                    }
                }
            }
        }

        Log-Debug "Calc UM count is: $UMCount"
        Log-Debug "Calc PersonalArchive count is: $PersonalArchiveCount"
        Log-Debug "Calc RetentionPolicy coutn is: $RetentionPolicyCount"
        Log-Debug "Calc ManagedFolder count is: $ManagedFolderCount"
        Log-Debug "Calc LegalHold count is: $LegalHoldCount"
    }

    Log-Exit $MyInvocation.MyCommand
}

#
# Calculate Enterprise CAL users for ActiveSync
#
function Calc-ECALForActiveSync
{
    Log-Enter $MyInvocation.MyCommand

    if (Process-Category "ActiveSync")
    {
        $ActiveSyncCount = 0

        Get-CASMailbox -ResultSize 'Unlimited' -Filter 'ActiveSyncEnabled -eq $true' | foreach {
            $CASMailbox = $_
            if (($CASMailbox.ActiveSyncMailboxPolicy -ne $null) -and $Script:ActiveSyncMailboxPolicyWithECALFeature.Contains($CASMailbox.ActiveSyncMailboxPolicy.ObjectGuid))
            {
                if ($Script:AllMailboxIDs.Contains($CASMailbox.Guid))
                {
                    if (Process-Mailbox $CASMailbox.Guid "ActiveSync")
                    {
                        $ActiveSyncCount++
                        $Script:EnterpriseCALMailboxIDs[$CASMailbox.Guid] = Get-MailboxOutputValue $CASMailbox
                    }
                }
            }
        }

        Log-Debug "Calc ActiveSyn count is: $ActiveSyncCount"
    }

    Log-Exit $MyInvocation.MyCommand
}


#
# Calculate Enterprise CAL users for Journaling
#
function Calc-ECALForJournaling
{
    Log-Enter $MyInvocation.MyCommand

    if (Process-Category "Journaling")
    {
        # Check all journaling mailboxes(include all version) for all journaling rules, and count current version mailbox as Enterprise CALs.
        foreach ($JournalRule in Get-JournalRule){
            # There are journal rules in the org.

            if ( $JournalRule.Recipient -eq $Null ) {
                Log-Debug "One journaling rule journals the whole org (all mailboxes require ECALs)"

                $Script:OrgWideJournalingEnabled = $True
                $Script:JournalingUserCount = $Script:AllVersionMailboxIDs.Count
                $Script:TotalEnterpriseCALs = $Script:TotalMailboxes

                break
            } else {
                $RecipientFilter = "((PrimarySmtpAddress -eq '" + $JournalRule.Recipient + "'))"
                Log-Info "RecipientFilter: $RecipientFilter"

                $JournalRecipient = Get-Recipient -Filter ($RecipientFilter)

                if ( $JournalRecipient -ne $Null ) {
                    # Note!!!
                    # Remote mailbox is NOT count here since it's totally different story.
                    if (($JournalRecipient.RecipientTypeDetails -eq 'UserMailbox') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'SharedMailbox') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'LinkedMailbox') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'MailContact') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'PublicFolder') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'LegacyMailbox') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'RoomMailbox') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'EquipmentMailbox') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'MailForestContact') -or
                        ($JournalRecipient.RecipientTypeDetails -eq 'MailUser')) {

                        # Journal a mailbox
                        if (Process-Mailbox $_.Guid "JournalingMailboxIDs")
                        {
                            $Script:JournalingMailboxIDs[$JournalRecipient.Guid] = $null
                        }
                    } elseif ( ($JournalRecipient.RecipientType -like "Mail*Group") -or ($JournalRecipient.RecipientType -like "Dynamic*Group") ) {
                        # Journal a DG or DDG.
                        # Get all mailbox members for the current journal DG/DDG and add to JournalingDGMailboxMemberIDs.
                        Get-JournalingGroupMailboxMember $JournalRecipient.Guid $JournalRecipient.RecipientType
                    }
                }
            }
        }

        if ( !$Script:OrgWideJournalingEnabled ) {
            # No journaling rules journaling the entire org.
            # Get all journaling mailboxes
            $Script:JournalingMailboxIDs = Merge-Hashtables $Script:JournalingDGMailboxMemberIDs $Script:JournalingMailboxIDs
            $Script:JournalingUserCount = $Script:JournalingMailboxIDs.Count

            # Calculate Enterprise CALs as not all mailboxes are Enterprise CALs
            foreach ($journalingMailboxID in $Script:JournalingMailboxIDs.Keys) {
                if ($Script:AllMailboxIDs.Contains($journalingMailboxID)) {
                    if (Process-Mailbox $journalingMailboxID "Journaling")
                    {
                        $Script:EnterpriseCALMailboxIDs[$journalingMailboxID] = $AllMailboxIDs[$journalingMailboxID].PrimarySmtpAddress
                    }
                }
            }
        }

        Log-Debug "Cacl Journaling count is: $($Script:JournalingUserCount)"
    }

    Log-Exit $MyInvocation.MyCommand
}

########################
## Script starts here ##
########################
Log-Info "The script will only query for the info and not change any settings."
Log-Debug "The script will run with `"-VersionMajor $Script:VersionMajor -AccessLicenseType $Script:AccessLicenseType -DebugCategory $Script:DebugCategory -DebugMailbox $Script:DebugMailbox`""

Set-DebugMailboxGuid

BackupSet-EnvironmentVariable

Set-ADServerSettings -ViewEntireForest $true

Calc-AllMaiboxes
if ($TotalMailboxes -eq 0)
{
    Log-Debug "No mailboxes in the org."
}
else
{
    # All users are counted as Standard CALs
    $Script:TotalStandardCALs = $Script:TotalMailboxes

    if ($Script:AccessLicenseType -ieq "Standard")
    {
        Log-Debug "Standard CALs were already calculated. All mailboxes require ECALs."
    }
    else
    {
        Log-Info "Calculating ECALs."

        if (Calc-ECALForOrg)
        {
            Log-Debug "ECALs per org were already calculated. All mailboxes require ECALs."
        }
        else
        {
            Log-Info "Calculating ECALs per mailbox."

            Cache-ForManagedFolder
            Cache-ForRetentionPolicy
            Cache-ForActiveSync

            Calc-ECALForMultiple
            Calc-ECALForActiveSync
            Calc-ECALForJournaling
        }
    }
}

$Script:TotalEnterpriseCALs = $Script:EnterpriseCALMailboxIDs.Count
Restore-EnvironmentVariable

Log-Info "Writing output objects"
Output-Report
# SIG # Begin signature block
# MIIdpAYJKoZIhvcNAQcCoIIdlTCCHZECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUTZtRSnOW/YlTZmTRfMr/cDgn
# i5SgghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkIxQjctRjY3Ri1GRUMyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEApkZzIcoArX4o
# w+UTmzOJxzgIkiUmrRH8nxQVgnNiYyXy7kx7X5moPKzmIIBX5ocSdQ/eegetpDxH
# sNeFhKBOl13fmCi+AFExanGCE0d7+8l79hdJSSTOF7ZNeUeETWOP47QlDKScLir2
# qLZ1xxx48MYAqbSO30y5xwb9cCr4jtAhHoOBZQycQKKUriomKVqMSp5bYUycVJ6w
# POqSJ3BeTuMnYuLgNkqc9eH9Wzfez10Bywp1zPze29i0g1TLe4MphlEQI0fBK3HM
# r5bOXHzKmsVcAMGPasrUkqfYr+u+FZu0qB3Ea4R8WHSwNmSP0oIs+Ay5LApWeh/o
# CYepBt8c1QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCaaBu+RdPA6CKfbWxTt3QcK
# IC8JMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAIl6HAYUhsO/7lN8D/8YoxYAbFTD0plm82rFs1Mff9WBX1Hz
# /PouqK/RjREf2rdEo3ACEE2whPaeNVeTg94mrJvjzziyQ4gry+VXS9ZSa1xtMBEC
# 76lRlsHigr9nq5oQIIQUqfL86uiYglJ1fAPe3FEkrW6ZeyG6oSos9WPEATTX5aAM
# SdQK3W4BC7EvaXFT8Y8Rw+XbDQt9LJSGTWcXedgoeuWg7lS8N3LxmovUdzhgU6+D
# ZJwyXr5XLp2l5nvx6Xo0d5EedEyqx0vn3GrheVrJWiDRM5vl9+OjuXrudZhSj9WI
# 4qu3Kqx+ioEpG9FwqQ8Ps2alWrWOvVy891W8+RAwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBKowggSmAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBvjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUNV1fNkMQ5b52/v6utPMXIcqK2rkwXgYKKwYB
# BAGCNwIBDDFQME6gJoAkAEMAYQBsAEMAYQBsAGMAdQBsAGEAdABpAG8AbgAuAHAA
# cwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZI
# hvcNAQEBBQAEggEAIfjwPnzV+GWEnBBdLnE8CovdrNIyaTy27/YWPh1zlKhnNaHf
# /WPF6QlvmuoAm5RQp/X5FYnDffb7PkoTUHP1vJKSOZie3M0Kw0LA2/lhJgseI5+k
# nzQU4f9lZX88XCudBxWrJBoejP2EyZ+0hRuUDuZpuL+SvAV3OuRf3NhnA9s/pC8a
# /oEDZbdvf3lFaOEmYHJUC+mqMfSVMiz7dbF9vSV4mNHP76kt2DmAFW7k6SZA5nhI
# IIR3ih/7z78ABzToMCBgR864TTiqROkmZSeNPY4FhKqNkLevC1FUIof4Y36DzNML
# +k8WX28DwZNhFVDByZ7kQwwa4D2N+OE7Auq0w6GCAigwggIkBgkqhkiG9w0BCQYx
# ggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACampsW
# woPa1cIAAAAAAJowCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0B
# BwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQ0MVowIwYJKoZIhvcNAQkEMRYE
# FHUD2lJViHv56GNj1XJx8s1Uh37mMA0GCSqGSIb3DQEBBQUABIIBAF+LRVmjPKna
# dNKmQGRHHzVJIb6kpQMRaSY2Be20kDH74GbfKyDNL0VsmUVte7lSY7mGz8Inzf3T
# 94SE4q/E0olp1u8thXRlBk3p6Y10rR7zH+tg5ST2Ggi2VNYuODjO862903ZKPDDd
# 98AdCvOd8x9bt9/SzrZMX3s3wtJUJ1lakTCleF/M1JNbvtRUkL3TpCayANqhGPVO
# atfnrhzSIPBUeldaYxmj3msH0xfBHOwb2ia4PxMeR0OFaX2cisCWltCJAjc4m6GO
# kmFRp28/qgfjTAJm16cWNoIFW9jJDzCQgL1doRk1Kt2q7MNBpA2HtA0xZ5P/sByp
# zLPHZnYlFbM=
# SIG # End signature block
