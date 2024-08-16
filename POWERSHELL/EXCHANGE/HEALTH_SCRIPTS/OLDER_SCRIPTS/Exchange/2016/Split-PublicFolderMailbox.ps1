# .SYNOPSIS
# Split-PublicFolderMailbox.ps1
#    Splits the given public folder mailbox based on the size of the folders
#
# .DESCRIPTION
#
# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Param(
    # Source Mailbox
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Please specify the source public folder mailbox to split")]
    [ValidateNotNull()]
    [string] $SourcePublicFolderMailbox,
    
    # Target Mailbox
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Please specify the target public folder mailbox where the split contents need to go")]
    [ValidateNotNull()]
    [string] $TargetPublicFolderMailbox,
    
    # Name of the organization
    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [string] $OrganizationName,
    
    # Minimum percentage to Split
    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [ValidateRange(1,100)]
    [int] $MinimumPercentageToSplit,
            
    # Source Public Folders Occupied Size
    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [int] $SourcePublicFoldersOccupied,
            
    # Move large items also
    [Parameter(Mandatory=$false)]
    [switch] $AllowLargeItems,
            
    [Parameter(Mandatory=$false)]
    [switch] $Whatif,

	# Don't move, just calculate the split plan
    [Parameter(Mandatory=$false)]
    [switch] $SplitPlan,

    # TEST HOOK!!!! Fake source mailbox quota
    [Parameter(Mandatory=$false)]
    [int] $FakeSourceMailboxQuota,

    # TEST HOOK!!!! Fake target mailbox quota
    [Parameter(Mandatory=$false)]
    [int] $FakeTargetMailboxQuota
   )

################ START OF DEFAULTS ################

# Folder Node's member indices
# This is an optimization since creating and storing objects as PSObject types
# is an expensive operation in powershell
# CLASSNAME_MEMBERNAME
$script:FOLDERNODE_PATH = 0;
$script:FOLDERNODE_TOTALITEMSIZE = 1;
$script:FOLDERNODE_AGGREGATETOTALITEMSIZE = 2;
$script:FOLDERNODE_PARENT = 3;
$script:FOLDERNODE_CHILDREN = 4;

# Default folder root node
$script:IPM_SUBTREE = "IPM_SUBTREE";
$script:customFolderPathSeparator = [char]255;

# Default minimum source percentage occupied which will warant a split (if more than 60% source total size is occupied, then the split is warranted)
$script:minSourcePercentOccupiedToSplit = 60;

# Default maximum target percent that can be occupied (including the moved content) in order to allow a split
# Try to maintain a 30% room for growth on target mailbox
$script:maxTargetPercentOccupiedToSplit = 70;

$script:minPercentQuotaSizeToFill = 10;
$script:maxPercentQuotaSizeToFill = 70;
$script:desiredPercentQuotaSizeToFill = 50;
$script:percentQuotaSizeToFillInStep1Sparse = 35;

# Passing in -Folders param of length > (256 *128) results in Xml ReaderQuota on MRS service to be exceeded, and move request would fail.
$script:maxLengthOfFolderNamesToMove = 256 * 128;
$script:currentLengthOfFolderNamesToMove = 0;
$script:AssignedFolders = @();
$script:UnAssignedFolders = @();
$script:PublicFoldersInfo = @{};
$script:FoldersToSplit = @();
$script:sourceMailboxInformation  = $null;
$script:targetMailboxInformation  = $null;

$script:ROOT = @($script:customFolderPathSeparator.ToString(), 0, 0, $null, @{});

# Script scoped variables to hold return values of functions that write to the pipeline using Write-Output
$script:EnsureSplitMailboxPossible = $null;
$script:CheckIfMoveContentPossible = $null;
$script:IsTransientError = $true;
$script:IsPermanentError = $false;

################ END OF DEFAULTS ################

############## Function to error out from the script execution ##########
function ErrorOut([string]$errorString, [System.Management.Automation.ErrorRecord] $lastErrorRecord, [bool] $isTransientError, [string] $errorReason)
{
    # If we need to return the split plan but we had an error, populate the ErrorRecord object with the error details before exiting
    if ($SplitPlan -eq $true)
    {
        $exceptionToReturn = $null;
        if ($lastErrorRecord -ne $null)
        {
            $exceptionToReturn = New-Object System.ApplicationException $errorString $lastErrorRecord.Exception
        }
        else
        {
            $exceptionToReturn = New-Object System.ApplicationException $errorString 
        }

        $exceptionToReturn.Data.Add("SplitPublicFolderMailbox:ErrorOut:ErrorReason", $errorReason)

        $errorID = 'PublicFolderSplitError'
        $errorCategory = [Management.Automation.ErrorCategory]::NotSpecified
        $target = $null
        $errorRecord = New-Object System.Management.Automation.ErrorRecord $exceptionToReturn, $errorID, $errorCategory, $target

        if ($isTransientError -eq $true)
        {
            throw $errorRecord
        }
        else
        {
            Write-Error -ErrorRecord $errorRecord
        }
    }

    exit;
}

############## Function to get information about quota ##########
function GetMailboxQuota()
{
    param ($mailboxType)
    
    if ($mailboxType -match "Source")
    {
        $useDatabaseQuotaDefault = $script:sourceMailboxInformation.UseDatabaseQuotaDefaults;
        $databaseName = $script:sourceMailboxInformation.Database.ToString();
        $prohibitSendReceiveQuota = $script:sourceMailboxInformation.ProhibitSendReceiveQuota.Value;
    }
    else
    {
        $useDatabaseQuotaDefault = $script:targetMailboxInformation.UseDatabaseQuotaDefaults;
        $databaseName = $script:targetMailboxInformation.Database.ToString(); 
        $prohibitSendReceiveQuota = $script:targetMailboxInformation.ProhibitSendReceiveQuota.Value;
    }
             
    if ($useDatabaseQuotaDefault -eq $true)
    {        
        if ($databaseName -ne $null)
        {            
            $prohibitSendReceiveQuota = (Get-MailboxDatabase -Identity:$databaseName).ProhibitSendReceiveQuota.Value; 
        }
        else
        {            
            return $null;
        }
    }    
    
    return $prohibitSendReceiveQuota.ToBytes();    
}
############## End Function #####################################

############## Function to ensure if mailbox can be split #######
function EnsureSplitMailboxPossible()
{
    # Determining ProhibitSendReceiveQuota value
    $sourceMailboxTotalSize = GetMailboxQuota("Source"); 
    $script:targetMailboxTotalSize = GetMailboxQuota("Target");

	# TEST HOOK!!!! Fake the source/destination quotas based on test parameters
	if ($FakeSourceMailboxQuota -ne $null -and $FakeSourceMailboxQuota -gt 0)
	{
		$sourceMailboxTotalSize = $FakeSourceMailboxQuota
	}

	if ($FakeTargetMailboxQuota -ne $null -and $FakeTargetMailboxQuota -gt 0)
	{
		$script:targetMailboxTotalSize = $FakeTargetMailboxQuota
	}
    
    if ($sourceMailboxTotalSize -eq $null)
    {
        Write-Output "";
        Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.UnlimitedSize -f $SourcePublicFolderMailbox);
        $script:EnsureSplitMailboxPossible = $false;
        return;
    }    
    
    # Retrieving statistics information on the source mailbox
    if ($OrganizationName -ne "")
    {
        $sourceMailboxOccupied = (Get-MailboxStatistics -Identity:($OrganizationName + "\" + $SourcePublicFolderMailbox)).TotalItemSize.Value.ToBytes();
    }
    else
    {
        $sourceMailboxOccupied = (Get-MailboxStatistics -Identity:$SourcePublicFolderMailbox).TotalItemSize.Value.ToBytes();
    }

    if ($sourceMailboxOccupied -eq $null)
    {
        ErrorOut "Failed to get source mailbox statistics" $error[0].Exception $script:IsTransientError "sourceMailboxOccupied";
    } 		
    
	# The split will attempt to move half of the size occupied by public folders alone. 
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.SourceStatistics -f $SourcePublicFolderMailbox, $sourceMailboxTotalSize, $sourceMailboxOccupied, $script:sourceMailboxPublicFoldersOccupied);

    # Folder sizes do not always make for a clean-cut split plan at 50% of the source occupied size
    # To accomodate various folder sizes in an optimal way, we will define a desired size to split, but allow non-optimal variations as well, in order to make progress
    # We will make it so we move at least the minimum size, ideally close to the desired size, but no more than max size
    $script:minQuotaSizeToFill = [Math]::Round(($script:sourceMailboxPublicFoldersOccupied * $script:minPercentQuotaSizeToFill / 100), 0);
    $script:maxQuotaSizeToFill = [Math]::Round(($script:sourceMailboxPublicFoldersOccupied * $script:maxPercentQuotaSizeToFill / 100), 0);
    $script:desiredQuotaSizeToFill = [Math]::Round(($script:sourceMailboxPublicFoldersOccupied * $script:desiredPercentQuotaSizeToFill / 100), 0);
    $script:quotaSizeToFillInStep1Sparse = [Math]::Round(($script:sourceMailboxPublicFoldersOccupied * $script:percentQuotaSizeToFillInStep1Sparse / 100), 0);
    
    # Determining if the minimum percentage is passed as an argument
    if ($MinimumPercentageToSplit -ne "")
    {
        $script:minSourcePercentOccupiedToSplit = [int]$MinimumPercentageToSplit;
    }
    
    # Computing percentage occupied; this uses the total items size, which includes content outside of the active folders (like items currently in retention after a previous move)
    $percentSourceMailboxOccupation = [Math]::Round(($sourceMailboxOccupied / $sourceMailboxTotalSize * 100), 0);
    if ($percentSourceMailboxOccupation -ge $script:minSourcePercentOccupiedToSplit)
    {
        $script:EnsureSplitMailboxPossible = $true;
        return;
    }
    else
    {        
        Write-Output "";
        Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.SplitSizeInformation -f $script:minSourcePercentOccupiedToSplit, $percentSourceMailboxOccupation);
        $script:EnsureSplitMailboxPossible = $false;
        return;
    }
}
############## End Function #####################################

#### Function to ensure if contents can be moved to the #########
################# given target mailbox ##########################
function CheckIfMoveContentPossible()
{
    # Retrieving statistics information on the target mailbox
    if ($OrganizationName -ne "")
    {
        $targetMailboxOccupied = (Get-MailboxStatistics -Identity:($OrganizationName + "\" + $TargetPublicFolderMailbox)).TotalItemSize.Value.ToBytes();
    }
    else
    {
        $targetMailboxOccupied = (Get-MailboxStatistics -Identity:$TargetPublicFolderMailbox).TotalItemSize.Value.ToBytes();
    }

    if ($targetMailboxOccupied -eq $null)
    {
        ErrorOut "Failed to get target mailbox statistics" $error[0].Exception $script:IsTransientError "targetMailboxOccupied";
    }
        
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.TargetStatistics -f $TargetPublicFolderMailbox, $script:targetMailboxTotalSize, $targetMailboxOccupied);

    $script:targetMailboxAvailableSize = [Math]::Round(($script:maxTargetPercentOccupiedToSplit / 100 * ($script:targetMailboxTotalSize - $targetMailboxOccupied)), 0);

    # Check to see if the given target mailbox is the right candidate
    if ($script:minQuotaSizeToFill -le $script:targetMailboxAvailableSize)
    {
        $script:CheckIfMoveContentPossible = $true;
        return;
    }
    else
    {
        $script:CheckIfMoveContentPossible = $false;
        return;
    }
}
############## End Function #####################################

########## Function that constructs the entire tree based on the folderpath ###############
###### As and when it constructs, it computes the aggregate folder size ###################
function LoadFolderHierarchy() 
{
    $publicFoldersToProcessCount = $script:publicFoldersToProcess.Count;
    for ($index = 0; $index -lt $publicFoldersToProcessCount; $index++)
    {
        $folderPaths = $script:publicFoldersToProcess[$index].FolderPath;
        $folderSize = $script:publicFoldersToProcess[$index].FolderSize;

        $folderName = "";   
        $parent = $script:ROOT;
        $jindex = 0;
        while ($folderPaths[$jindex] -ne $null)
        {
            $folderName = $folderName + $script:customFolderPathSeparator + $folderPaths[$jindex];
            $child = $parent[$script:FOLDERNODE_CHILDREN][$folderName];
            if ($child -eq $null)
            {
                # Construct a default node
                $child = @($folderName, 0, 0, $parent, @{});
                
                # Populate the node if it is not going to be an internal node
                if ($folderPaths[$jindex + 1] -eq $null)
                {
                    $child[$script:FOLDERNODE_TOTALITEMSIZE] = $folderSize;
                    $child[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] = $folderSize;
                    $script:PublicFoldersInfo[$folderName] = $script:publicFoldersToProcess[$index];
                }
                
                $parent[$script:FOLDERNODE_CHILDREN][$folderName] = $child;
            }            
            
            $parent[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] = $parent[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] + $folderSize;
            $parent = $child;
            $jindex++;
        }                
    } 

	# Print out the folder hierarchy  "," $folder[$script:FOLDERNODE_TOTALITEMSIZE] "," $folder[$script:FOLDERNODE_AGGREGASTEDTOTALITEMSIZE] "," $folder[$script:FOLDERNODE_PARENT][$script:FOLDERNODE_PATH] "," @($folder[$script:FOLDERNODE_CHILDREN].Values).Count);
	Write-Output "";Write-Output "";
	Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.FoldersInHierarchy;
	Write-Output "Path, FolderSize, AggregatedSize";
	PrintMailboxPublicFolderHierarchy $script:ROOT;
} 
############## End Function ###############################################################

############# Function that assigns Public Folders content to the target mailbox, up to the desired quota size to fill ##########################
############# Viable candidates are stored in $script:AssignedFolders. Remaining candidates are kept in a separate list $script:UnAssignedFolders ##########################
############# $node: Node to be assigned to the mailbox #######################################
function AllocateMailbox()
{
    param ($node, $partOfSparseBranch)

	if ($partOfSparseBranch -eq $true)
	{
		# Once we decided a branch is sparse, we will subject the entire sub-tree to the sparse quota, just so we don't fill the target 
		# with a lot of small isolated items that have no chance of being consolidated with the parents
		$quotaSizeToCheck = $script:quotaSizeToFillInStep1Sparse;
	}
	else
	{
		$quotaSizeToCheck = $script:desiredQuotaSizeToFill;
	}

    if (($node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] -le $script:targetMailboxAvailableSize) -and ($node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] + $script:targetFilledSize -le $quotaSizeToCheck))
    {
		$folderObject = New-Object PSObject -Property @{FolderPath = $node[$script:FOLDERNODE_PATH]; FolderSize = $node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE]; Aggregated = $true; SelectedForSplit = $false; SelectedInPass = "0"};
		if ($node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] -eq 0)
		{
			# Empty leaf folder. We will not allocate these unless they are children of allocated parents
			$script:UnAssignedFolders += $folderObject;
		}
		else
		{
			# Node's contents (including branch) can be completely fit into target mailbox
			# Assign the folder to mailbox and update mailbox's remaining size
			$script:targetMailboxAvailableSize -= $node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE];
			$script:targetFilledSize += $node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE];
			$script:AssignedFolders += $folderObject;
			$folderObject.SelectedInPass = "1";
		}
    }
    else
    {
        # Since node's contents (including branch) could not be fitted into the target mailbox, see if we can put it's individual contents into the mailbox
		# We have a lower threshold for allowing sparse branches in step 1, to increase the likelihood that we can accomodate the entire branch in step 2 or 3.
		$folderObject = New-Object PSObject -Property @{FolderPath = $node[$script:FOLDERNODE_PATH]; FolderSize = $node[$script:FOLDERNODE_TOTALITEMSIZE]; Aggregated = $false; SelectedForSplit = $false; SelectedInPass = "0"};
		if (($node[$script:FOLDERNODE_TOTALITEMSIZE] -gt 0) -and ($node[$script:FOLDERNODE_TOTALITEMSIZE] -le $script:targetMailboxAvailableSize) -and ($node[$script:FOLDERNODE_TOTALITEMSIZE] + $script:targetFilledSize -le $script:quotaSizeToFillInStep1Sparse))
		{
			# Node's contents (but not branch) can be completely fit into target mailbox
			# Assign the folder to mailbox and update mailbox's remaining size
			$script:targetMailboxAvailableSize -= $node[$script:FOLDERNODE_TOTALITEMSIZE];
			$script:targetFilledSize += $node[$script:FOLDERNODE_TOTALITEMSIZE];
			$script:AssignedFolders += $folderObject;
			$folderObject.SelectedInPass = "1";
			$partOfSparseBranch	= $true;
		}
		else
		{
			# This folder's individual content is either empty or it doesn't fit. Keep it unassigned for now
			$script:UnAssignedFolders += $folderObject;
		}

		# Whether or not we could allocate the folder's individual content, try to allocate the subfolders at this point
		$subFolders = @(@($node[$script:FOLDERNODE_CHILDREN].Values) | Sort-Object @{Expression={$_[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE]}; Ascending=$true});
		foreach ($subFolder in $subFolders)
		{
			AllocateMailbox $subFolder $partOfSparseBranch;
		}
	}
}
############## End Function #####################################################################

############# Function that prints the Public Folders hierarchy ##########################
############# $node: Node to print #######################################
function PrintMailboxPublicFolderHierarchy()
{
    param ($node)
	$folderObject = New-Object PSObject -Property @{FolderPath = $node[$script:FOLDERNODE_PATH]; FolderSize = $node[$script:FOLDERNODE_TOTALITEMSIZE]; AggregatedSize = $node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE]};
	Write-Output $folderObject | ft -HideTableHeaders -Autosize -Wrap FolderPath, FolderSize, AggregatedSize;
	$subFolders = @(@($node[$script:FOLDERNODE_CHILDREN].Values) | Sort-Object @{Expression={$_[$script:FOLDERNODE_PATH]}; Ascending=$true});
	foreach ($subFolder in $subFolders)
	{
		PrintMailboxPublicFolderHierarchy $subFolder;
	}
}
############## End Function #####################################################################

############## Function to try assigning children with parents ####
### We will try to accomodate subfolders with parent in the condition below:
###         - We haven't moved enough content (we currently have assigned less than desiredQuotaSizeToFill) 
###			- Adding the subfolder will not cause us to go over maxQuotaSizeToFill
function TryAccomodateUnassignedSubFoldersWithAssignedParent()
{
    $numUnAssignedFolders = $script:UnAssignedFolders.Count;
    for ($index = $numUnAssignedFolders - 1 ; $index -ge 0 ; $index--)
    {
		# If we already filled the desired size, we are done
		if ($script:targetFilledSize -gt $script:desiredQuotaSizeToFill)
		{
			break;
		}

        # See if we can accomodate this folder based on its size
		$unAssignedFolder = $script:UnAssignedFolders[$index];
		if ($unAssignedFolder.FolderSize -gt $script:targetMailboxAvailableSize -or $unAssignedFolder.FolderSize + $script:targetFilledSize -gt $script:maxQuotaSizeToFill)
		{
			# This folder is too large
			continue;
		}

        # Try to locate folder's parent in the assigned folders list
	    $numAssignedFolders = $script:AssignedFolders.Count;
        for ($jindex = $numAssignedFolders - 1 ; $jindex -ge 0 ; $jindex--)
        {
            if ($unAssignedFolder.FolderPath.StartsWith($script:AssignedFolders[$jindex].FolderPath))
            {
                # Found an ancestor, so select the child as well. 
				# This will be done by moving the folder node from the unassigned list to the assigned list
                $ancestor = $script:AssignedFolders[$jindex];
				$script:AssignedFolders += $unAssignedFolder;
				$script:targetFilledSize += $unAssignedFolder.FolderSize;
				$script:targetMailboxAvailableSize -= $unAssignedFolder.FolderSize;
                $script:UnAssignedFolders[$index] = $null;
				$unAssignedFolder.SelectedInPass = "2";

                break;
            }
        }
    }
    
    if ($script:UnAssignedFolders.Count -gt 1)
    {
        $script:UnAssignedFolders = $script:UnAssignedFolders | where {$_ -ne $null};
    }
}
############## End Function #####################################################################

############## Function to try assigning parents with children ####
### We will try to accomodate parents with already allocated subfolders in the condition below:
###         - We haven't moved enough content (we currently have assigned less than desiredQuotaSizeToFill) 
###			- Adding the subfolder will not cause us to go over maxQuotaSizeToFill
###         - All
function TryAccomodateUnassignedParentWithAssignedSubfolders()
{
    $numUnAssignedFolders = $script:UnAssignedFolders.Count;
    $numAssignedFolders = $script:AssignedFolders.Count;
    for ($index = $numUnAssignedFolders - 1 ; $index -ge 0 ; $index--)
    {
		# If we already filled the desired size, we are done
		if ($script:targetFilledSize -gt $script:desiredQuotaSizeToFill)
		{
			break;
		}

        # See if we can accomodate this folder based on its size
		$unAssignedFolder = $script:UnAssignedFolders[$index];
		if ($unAssignedFolder.FolderSize -gt $script:targetMailboxAvailableSize -or $unAssignedFolder.FolderSize + $script:targetFilledSize -gt $script:maxQuotaSizeToFill)
		{
			# This folder is too large
			continue;
		}

        # Try to locate folder's children in the assigned folders list
        for ($jindex = 0 ; $jindex -lt $numAssignedFolders ; $jindex++)
        {
            if ($script:AssignedFolders[$jindex].FolderPath.StartsWith($unAssignedFolder.FolderPath))
            {
                # Found an allocated child
				# See if we need to select the parent as well. We only select a parent if all children are selected
				$unselectedChildren = $false;
			    for ($workIndex = $numUnAssignedFolders - 1 ; $workIndex -gt $index ; $workIndex--)
				{
					if ($script:UnAssignedFolders[$workIndex] -ne $null -and $script:UnAssignedFolders[$workIndex].FolderPath.StartsWith($unAssignedFolder.FolderPath))
					{
						$unselectedChildren = $true;
						break;
					}
				}

				if ($unselectedChildren -eq $false)
				{
					# There are no unselected children, so we can move the parent
					# This will be done by moving the folder node from the unassigned list to the assigned list
					$child = $script:AssignedFolders[$jindex];
					$script:AssignedFolders += $unAssignedFolder;
					$script:targetFilledSize += $unAssignedFolder.FolderSize;
					$script:targetMailboxAvailableSize -= $unAssignedFolder.FolderSize;
					$script:UnAssignedFolders[$index] = $null;
					$unAssignedFolder.SelectedInPass = "3";
				}

                break;
            }
        }
    }
    
    if ($script:UnAssignedFolders.Count -gt 1)
    {
        $script:UnAssignedFolders = $script:UnAssignedFolders | where {$_ -ne $null};
    }
}
############## End Function #####################################################################

# There are times when both parent and children be part of the identified folders for split
# This function aids the EnumerateFoldersToSplit function in making sure such children are not 
# part of the output when its parent is enumerated
function IsFolderPresent()
{
    param ($folderPath)
    
    $numAssignedFolders = $script:AssignedFolders.Count;
    for ($index = $numAssignedFolders - 1 ; $index -ge 0 ; $index--)
    {
        if ($script:AssignedFolders[$index].FolderPath -eq $folderPath)
        {
            return $true;
        }
    }
    
    return $false;
}
############## End Function #####################################################################

#################### Function to enumerate folders that are part of split job ##################
####### It enumerates by looking at all the folders that match the prefix of passed in folder###
function EnumerateFoldersToSplit()
{
    param ($node)
    
    $folderPath = "";    
    if ($node.FolderPath -ne $script:customFolderPathSeparator.ToString())
    {
        $folderPath = $node.FolderPath;    
    }

	if ($node.Aggregated -eq $true)
	{
	    # Walk through the folders to identify the matching children and itself
		foreach ($key in $script:PublicFoldersInfo.Keys)
		{
			$folderIdentity = $null;
			if ($key -eq $folderPath)
			{
				$folderIdentity = $script:PublicFoldersInfo[$key].Identity;            
			}
			# Escape special regex characters in folderPath
			elseif ($key -match ("^"+[regex]::escape($folderPath + $script:customFolderPathSeparator.ToString())))
			{
				$retVal = IsFolderPresent($key);
				if ($retVal -ne $true)
				{
					$folderIdentity = $script:PublicFoldersInfo[$key].Identity;                
				}
			}

			# If we found a matching folder, check if we would exceed the max length for -Folders param by adding it.
			# If yes, stop enumerating folders and return.
			if($folderIdentity -ne $null)
			{
				if($script:currentLengthOfFolderNamesToMove + $folderIdentity.ToString().Length -gt $script:maxLengthOfFolderNamesToMove)
				{               
					return $false;
				}
            
				$script:currentLengthOfFolderNamesToMove += $folderIdentity.ToString().Length;
				$script:FoldersToSplit += $folderIdentity;
			}
		}
	}
	else
	{
	    # Since this folder is not aggregated, enumeration should only pick the folder itself
		$folderIdentity = $script:PublicFoldersInfo[$folderPath].Identity;            

		# Check if we would exceed the max length for -Folders param by adding it. If yes, stop enumerating folders and return.
		if($folderIdentity -ne $null)
		{
			if($script:currentLengthOfFolderNamesToMove + $folderIdentity.ToString().Length -gt $script:maxLengthOfFolderNamesToMove)
			{               
				return $false;
			}
            
			$script:currentLengthOfFolderNamesToMove += $folderIdentity.ToString().Length;
			$script:FoldersToSplit += $folderIdentity;
		}
	}

    return $true;
}
############## End Function #####################################################################


#################### Function to select multiple folders for the move job ##################
####### It scans at the candidates for split (AssignedFolders) in the order in which they were enumerated #########
####### and selects them for split, if adding them to the list of folders to be moved #############################
####### does not cause the total length of folder names to exceed the maxLengthOfFolderNamesToMove. ###############
####### For any un-expanded branch, it also does the branch expansion into individual folders. ####################
function SelectFoldersToMove()
{
    for($i = 0 ; $i -lt $script:AssignedFolders.Count ; $i++)
    {    
        # Mark this branch as selected and enumerate its folders
        $script:AssignedFolders[$i].SelectedForSplit = $true;
        $enumerateComplete = EnumerateFoldersToSplit($script:AssignedFolders[$i]);
        
        # Enumerating folders would halt if it hits max folder name length for the -Folders param of New-PublicFolderMoveRequest.
        # In that case, stop selecting more branches and return.  
        if($enumerateComplete -eq $false)
        {
            break;
        }            
    }    
}
############## End Function #####################################################################

#load hashtable of localized string
Import-LocalizedData -BindingVariable PublicFolderManagement_LocalizedStrings -FileName PublicFolderStorageManagement.strings.psd1

if ($SourcePublicFolderMailbox -eq $TargetPublicFolderMailbox)
{
    Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.SameMailbox;
    ErrorOut $PublicFolderManagement_LocalizedStrings.SameMailbox $null $script:IsPermanentError "sourceEqualsTarget";
}

# Retrieve the information about the given source public folder mailbox
if ($OrganizationName -ne "")
{
    $script:sourceMailboxInformation = Get-Mailbox -PublicFolder -Identity:$SourcePublicFolderMailbox -Organization:$OrganizationName;
}
else
{
    $script:sourceMailboxInformation = Get-Mailbox -PublicFolder -Identity:$SourcePublicFolderMailbox;
}

if ($script:sourceMailboxInformation -eq $null)
{
        ErrorOut "Failed to get source mailbox information" $error[0].Exception $script:IsTransientError "sourceMailboxInformation";
}
    
# Retrieve the information about the given target public folder mailbox
if ($OrganizationName -ne "")
{
    $script:targetMailboxInformation = Get-Mailbox -PublicFolder -Identity:$TargetPublicFolderMailbox -Organization:$OrganizationName;
}
else
{
    $script:targetMailboxInformation = Get-Mailbox -PublicFolder -Identity:$TargetPublicFolderMailbox;
}

if ($script:targetMailboxInformation -eq $null)
{
        ErrorOut "Failed to get target mailbox information" $error[0].Exception $script:IsTransientError "targetMailboxInformation";
}

# Get the folders that belong to input source mailbox
Write-Output "";
Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.RetrieveFoldersFromSourceMailbox;
if ($OrganizationName -ne "")
{ 
    $script:publicFoldersToProcess = Get-PublicFolder -Organization:$OrganizationName -ResidentFolders -Mailbox $script:sourceMailboxInformation.ExchangeGuid.ToString() -Recurse -WarningAction SilentlyContinue | `
                                     Where-Object -FilterScript { $_.Name -ne $script:IPM_SUBTREE };
}
else
{
    $script:publicFoldersToProcess = Get-PublicFolder -ResidentFolders -Mailbox $script:sourceMailboxInformation.ExchangeGuid.ToString() -Recurse -WarningAction SilentlyContinue | `
                                     Where-Object -FilterScript { $_.Name -ne $script:IPM_SUBTREE };
}

# Check if there are atleast two folders that reside in the given source public folder mailbox to split
if ($script:publicFoldersToProcess.Count -lt 2)
{
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.NotEnoughFoldersToSplit -f $SourcePublicFolderMailbox);
    ErrorOut ($PublicFolderManagement_LocalizedStrings.NotEnoughFoldersToSplit -f $SourcePublicFolderMailbox) $null $script:IsPermanentError "publicFoldersToProcess";
}

# If we don't have the source public folders occupied size, we need to load the folder hierarchy here in order to calculate correct split quotas to fill
# The TotalItemSize from mailbox statistics can be incorrect, as it includes content currently in retention
if ($SourcePublicFoldersOccupied -eq 0)
{
	LoadFolderHierarchy;
	$script:sourceMailboxPublicFoldersOccupied = $script:ROOT[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE];
}
else
{
	$script:sourceMailboxPublicFoldersOccupied = $SourcePublicFoldersOccupied;
}

# Check if is feasible to split the mailbox at this point
Write-Output "";
Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.FeasibilityToSplit;
EnsureSplitMailboxPossible;
if ($script:EnsureSplitMailboxPossible -ne $true)
{
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.ImpossibleToSplit -f $SourcePublicFolderMailbox);
    ErrorOut ($PublicFolderManagement_LocalizedStrings.ImpossibleToSplit -f $SourcePublicFolderMailbox) $null $script:IsPermanentError "EnsureSplitMailboxPossible";
}

# Check if is feasible to move contents to the given target mailbox
Write-Output "";
Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.FeasibilityToMove;
CheckIfMoveContentPossible;
if ($script:CheckIfMoveContentPossible -ne $true)
{
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.ImpossibleToMove -f $TargetPublicFolderMailbox);
    ErrorOut ($PublicFolderManagement_LocalizedStrings.ImpossibleToMove -f $TargetPublicFolderMailbox) $null $script:IsPermanentError "CheckIfMoveContentPossible";
}

# Load the folder hierarchy if we haven't already
if ($SourcePublicFoldersOccupied -ne 0)
{
	LoadFolderHierarchy;
}

Write-Output "";
Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.IdentifyFolders;

# Allocate folders up to desired quota
$script:targetMailboxFilledSize = 0;
AllocateMailbox $script:ROOT $false;

# If we haven't reached desired quota, try picking some subfolders of already allocated parents
TryAccomodateUnassignedSubFoldersWithAssignedParent;

# If we haven't reached desired quota, try picking some parents of already allocated children
TryAccomodateUnassignedParentWithAssignedSubfolders;

# Print out the split candidates
if ($script:AssignedFolders.Count -ge 1)
{
    Write-Output "";Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.CandidatesForSplit -f $script:targetFilledSize);
    $script:AssignedFolders | Select-object @{Name="FolderPath"; Expression={$_.FolderPath -Replace $script:customFolderPathSeparator,"\"}}, FolderSize, SelectedInPass;
                              
    Write-Output "";Write-Output "";
}
else
{
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.NotEnoughFoldersToSplit -f $SourcePublicFolderMailbox);
    ErrorOut ($PublicFolderManagement_LocalizedStrings.NotEnoughFoldersToSplit -f $SourcePublicFolderMailbox) $null $script:IsPermanentError "TryAccomodateUnassignedParentWithAssignedSubfolders";
}

# Select the folders to move
SelectFoldersToMove;

# Print out the folders selected for move
if ($script:AssignedFolders.Count -ge 1)
{
    Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.SelectedForSplit;
    Write-Output "";
    $script:AssignedFolders | Where-Object -FilterScript { $_.FolderPath -ne $script:customFolderPathSeparator.ToString() -and $_.SelectedForSplit -eq $true } | `
                              Select-object @{Name="FolderPath"; Expression={$_.FolderPath -Replace $script:customFolderPathSeparator,"\"}}, FolderSize;
                              
    Write-Output "";Write-Output "";
}
else
{
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.NotEnoughFoldersToSplit -f $SourcePublicFolderMailbox);
    ErrorOut ($PublicFolderManagement_LocalizedStrings.NotEnoughFoldersToSplit -f $SourcePublicFolderMailbox) $null $script:IsPermanentError "SelectFoldersToMove";
}

Write-Output "";
Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.MoveFolders;
$script:FoldersToSplit | Format-Wide -Property MapiFolderPath -Column 1;

if ($Whatif)
{
    exit;
}

# If we need to return the split plan, populate the split plan object before exiting
if ($SplitPlan -eq $true)
{
	$publicFoldersInfoByIdentity = @{}
	foreach ($folderName in $script:PublicFoldersInfo.Keys)
	{
		$publicFoldersInfoByIdentity[$script:PublicFoldersInfo[$folderName].Identity] = $script:PublicFoldersInfo[$folderName];
	}

	#load binary dependencies
	$exchangeInstallPath=(get-itemproperty hklm:\software\microsoft\exchangeServer\v15\setup).MsiInstallPath
	$exchangeInstallPath = $exchangeInstallPath.trimEnd("\")
	$dataStorageModule = $exchangeInstallPath + "\bin\Microsoft.Exchange.Data.Storage.dll"
	[System.Reflection.Assembly]::LoadFrom($dataStorageModule);

	# Prepare the split plan object
	$splitPlanFolders = @();
	foreach ($folderIdentity in $script:FoldersToSplit)
	{
		$splitFolderObject = New-Object Microsoft.Exchange.Data.Storage.PublicFolder.SplitPlanFolder;
		$splitFolderObject.PublicFolderId = $folderIdentity;
		$splitFolderObject.ContentSize = $publicFoldersInfoByIdentity[$folderIdentity].FolderSize;
        $splitPlanFolders += $splitFolderObject;
	}

    $splitPlanObject = New-Object Microsoft.Exchange.Data.Storage.PublicFolder.PublicFolderSplitPlan;
    $splitPlanObject.FoldersToSplit = $splitPlanFolders;
    $splitPlanObject.TotalSizeToSplit = $script:targetFilledSize;
    $splitPlanObject.TotalSizeOccupied = $script:sourceMailboxPublicFoldersOccupied;
    Write-Output $splitPlanObject;
    exit;
}

# Checking if there are any pending request in this organization
if ($OrganizationName -ne "")
{
    $anyPendingRequest = Get-PublicFolderMoveRequest -Organization:$OrganizationName;
}
else
{
    $anyPendingRequest = Get-PublicFolderMoveRequest;
}
    
if ($anyPendingRequest -ne $null)
{
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.RemoveExistingRequest;
    exit;
}

# Initiating the request
Write-Output "";
Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.IssueSplitRequest -f $SourcePublicFolderMailbox);
if ($OrganizationName -ne "")
{
    $request = New-PublicFolderMoveRequest -Folders:$script:FoldersToSplit -TargetMailbox:$TargetPublicFolderMailbox -AllowLargeItems:$AllowLargeItems -Organization:$OrganizationName;
}
else
{
    $request = New-PublicFolderMoveRequest -Folders:$script:FoldersToSplit -TargetMailbox:$TargetPublicFolderMailbox -AllowLargeItems:$AllowLargeItems;
}

if ($request -ne $null)
{
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.RequestName -f $($request));
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.SourceMailbox -f $($request.SourceMailbox));
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.TargetMailbox -f $($request.TargetMailbox));
    Write-Output "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.RequestStatus -f $($request.Status));
    Write-Output "";
    Write-Output "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.JobStatus;
}
# SIG # Begin signature block
# MIIdugYJKoZIhvcNAQcCoIIdqzCCHacCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUQEJZe8WOw7QEiOquzWYfnMtb
# 4jegghhkMIIEwzCCA6ugAwIBAgITMwAAAK7sP622i7kt0gAAAAAArjANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBMAwggS8AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB1DAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUw7lSDxDSamCnJHkBogO+/0xcfKkwdAYKKwYB
# BAGCNwIBDDFmMGSgPIA6AFMAcABsAGkAdAAtAFAAdQBiAGwAaQBjAEYAbwBsAGQA
# ZQByAE0AYQBpAGwAYgBvAHgALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAGPGYbDg2gppD6gJnSDA
# jOSdesxNjmt2oi88QYPqpvdqLuhtQQ5r1N4/iQptkvemN+akJ6I2yK5a+wKvDXVH
# fqIDGbmSWwG/zX+e+z7HHJZ0K3cu08xIyoDL/RP5jpJar7QYZUG6v20+fIaOaG3Y
# xm4JNUJnD3OjiBnb53kQv1y6R6rKoZ9n3qq5A5JzsSkZ7/8mFcT9rAc00Q8vK8AJ
# J+OjmvW2uMQq8SvEAA3VLr1ADCUnGPlFtfRkQ2qXF0aVpMJNw6pHzemeyWzFqsJX
# 5B/zsX4s2ppsdVs/sc6aaE6kasFWSGiF/nK2HuCxJKWSEavCUZ4gEQXHZjAmYQcy
# qx2hggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBAhMzAAAAruw/rbaLuS3SAAAAAACuMAkGBSsOAwIaBQCgXTAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMx
# ODQzNTFaMCMGCSqGSIb3DQEJBDEWBBTfzGLlBr/Ihc3klMKX1fJjePAWqDANBgkq
# hkiG9w0BAQUFAASCAQBtvTTQm/ZE5vl3Z7barfGERi6BWRvD+YuosMJHFpE8+p7M
# lo/1BMCUE+J4mnWY5ZB0CEKa/s8RnUpAE7Moh0L3C5rIT3Cyo1zEyOR3qQpN4xWR
# 1FZsLxpYNiO6RolJhbd+f9lAzYcd2gi9tNlbg5FrWc55NNlPMoO9SqRtSx6BzTrg
# or7XQkGzUe3hSB0y2SmXhGZWrohnCEcOP8wMYuHxFfhLLI9SRhkQ1y9ph5SPq6dZ
# tywVvKYBmef9Sr0SdvzOxTIxR0Kya2r7ZozP+T4LuZcFFWT6aGGt+1Iq/St7quC9
# SkQk7BdMRMvM7AXY6DAYt19uPJbJNyYdCkgSeU36
# SIG # End signature block
