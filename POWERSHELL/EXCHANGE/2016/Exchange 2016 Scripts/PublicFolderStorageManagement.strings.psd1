# Localized	09/03/2016 06:56 AM (GMT)	303:4.80.0411 	PublicFolderStorageManagement.strings.psd1
ConvertFrom-StringData @'
###PSLOC
SameMailbox=Please provide a target mailbox that is different from the source mailbox.
UnlimitedSize=Mailbox {0} has unlimited size.
FeasibilityToSplit=Checking if it is possible to split the given source mailbox
FeasibilityToMove=Checking if it is possible to move contents to the given target mailbox
FeasibilityToMerge=Checking if it is possible to merge the given source mailbox
SplitSizeInformation=Minimum percentage to split is {0} while percentage occupied by source mailbox is {1}.
ImpossibleToSplit=Public folder mailbox {0} cannot be split at this point.
ImpossibleToMove=Public folder mailbox {0} is not the right candidate to accomodate the moving contents.
ImpossibleToMerge=Public folder mailbox {0} cannot be merged at this point.
RetrieveFoldersFromSourceMailbox=Determining folders that belong to source mailbox
NotEnoughFoldersToSplit=There aren't enough folders residing in the mailbox {0} to split.
FolderUnavailableToMerge=There isn't any folder residing in the mailbox {0} to merge.
IdentifyFolders=Identifying folders that are to be moved to the target mailbox
FoldersInHierarchy=Folders in the public folder hierarchy on the source mailbox
CandidatesForSplit=Possible folder branches for splitting, with total size: {0}
SelectedForSplit=Selected folder branches for splitting
MoveFolders=Folders that will be moved as part of this request:
RemoveExistingRequest=Please remove the existing request and then continue...
IssueSplitRequest=Issuing request to split the mailbox: {0}
IssueMergeRequest=Issuing request to merge the mailbox: {0}
IssueMoveBranchRequest=Issuing request to move the public folder branch: {0}
RequestName=RequestName: {0}
SourceMailbox=SourceMailbox: {0}
TargetMailbox=TargetMailbox: {0}
RequestStatus=RequestStatus: {0}
JobStatus=Use Get-PublicFolderMoveRequest cmdlet to obtain the status of the job.
SourceStatistics=Source mailbox statistics: Mailbox: {0} MailboxSize: {1} OccupiedSize: {2} PublicFoldersOccupiedSize: {3}
TargetStatistics=Target mailbox statistics: Mailbox: {0} MailboxSize: {1} OccupiedSize: {2}
###PSLOC
'@
