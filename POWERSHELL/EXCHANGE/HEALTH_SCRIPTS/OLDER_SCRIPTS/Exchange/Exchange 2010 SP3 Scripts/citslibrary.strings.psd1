#####################################################################################
#
#
# THIS FILE EXISTS IN TWO LOCATIONS. MAKE SURE TO BOTH COPIES OF THE FILE ARE UPDATED WHEN 
# WHEN EITHER COPY IS CHANGED.
# <DEPOT>\Sources\dev\management\src\management\scripts\troubleshooter\CITSLibrary.strings.psd1
# <DEPOT>\Sources\dev\mgmtpack\src\HealthMainfests\scripts\troubleshooter\CITSLibrary.strings.psd1
# The management version of the library gets deployed during exchange setup and the 
# mgmtpack version of the library only gets deployed when the management pack is installed
#
######################################################################################

ConvertFrom-StringData @'
###PSLOC
    # English strings
    AllNotAllowedForResolve = When 'Resolve' is specified for the Action parameter, 'All' is not allowed for the Symptom parameter.
    TroubleshooterFailed = The troubleshooter failed with error: 
    RegistryOpenError = Could not open the registry on server: 
    RegistryReadError = Couldn't read registry key: 
    TimeoutWaitingForEvent = Event couldn't be found within the timeout period:
    TimeoutWaitingForProcessToStop = The process didn't stop within the time-out period:
    TimeoutZeroOrNegative = Timeout reached zero or negative value entering function:

# Event Log strings 
# Logged both to application and crimson logs   
#
    TSStarted=The troubleshooter started successfully. Version: %1
    DetectedDeadlock=Detected search service deadlock.
    DetectedCatalogCorruption=Detected catalog corruption for database %1
    DetectedIndexingStall=Detected indexing stall for database %1. Stall counter value %2. Stall threshold value is %3 seconds
    DetectedIndexingStallExtendedPeriod=Detected indexing stall for database %1 for an extended duration. Stall counter value %2. Stall threshold value is %3 seconds
    DetectedIndexingBacklog=Indexing backlog reached a critical limit of %2 hours or the number of items in the retry queue is greater than %3 for database: %1
    DetectedIndexingBacklogOrLargeRetryQueuesOnMultipleDatabases=Indexing backlog reached a critical limit of %2 hours or the number of items in the retry queue is greater than %3 for one or more databases: %1
    TSFailed=The troubleshooter failed with exception %1.
    TSSuccess=The troubleshooter finished successfully.
    RestartSuccess=Restart of search services succeeded.
    RestartFailure=Search services failed to restart. Reason: %1. Current server status: %2
    DetectedNoIssues=The troubleshooter didn't find any issues for any catalog.
    CatalogHasNoIssues=The troubleshooter didn't find any catalog issues for database %1.
    DetectedSameSymptomTooManyTimes=The troubleshooter detected the symptom %1 %2 times in the past %3 hours for catalog %4. This exceeded the allowed limit for failures.
    ReseedSuccess=Reseeding succeeded for the catalog of database %1.
    ReseedFailure=Reseeding failed for the content index catalog of mailbox database %1. Reason: %2. Database copy states: %3
    ActiveCatalogCopyCorrupt=The active catalog of mailbox database '%1' is corrupt. Database copy states: %2
    AnotherInstanceRunning=Another instance of the troubleshooter is already running on this machine. Two or more instances cannot be run simultaneously.
    MsftefdMemoryUsageHigh=The memory usage of the Msftefd processes is above the set limit of %1 Mb. Current value %2 Mb. Process instances %3
    MsftefdMemoryUsageHighWithCrashDump=The memory usage of the Msftefd processes is above the set limit of %1 Mb. Current value %2 Mb. A crash dump of the process has also been taken.
    FoundBadIFiltersEnabled=The troubleshooter found the following %1 Filters which should be disabled.
    EnablingIFilterFailed=The troubleshooter failed to enable IFilter %1. Reason %2
    IFiltersToEnable=The troubleshooter found the following IFilters to enable:  %1
    CatalogSizeGreaterThanExpectedDBLimit=The percentage catalog size for mailbox database %1 is greater than the threshold value of %2. Current value %3
    CatalogReseedLoop=The catalog for mailbox database %1 has been reseeded %2 consecutive times by the CI troubleshooter.
    TroubleshooterDisabled=The troubleshooter has been disabled on server %1
    ItemsStuckInRetryQueue=The search service is not processing items in the retry queue for mailbox database %1. Documents Processed since last run %2.
    ServiceRestartNotNeeded=The CI troubleshooter did not detect any issue that would require a restart of the search service.
    ServiceRestartAttempt=CI troubleshooter exchange search service restart attempt %1.
    CatalogRecoveryDisabled=Recovery actions for search catalog %1 has been disabled.
    RetryQueuesStagnant=The search service is not processing items in the retry queue for mailbox databases [CatalogName (AgeOfLastNotificaton, NumberOfItemsInRetryQueue, NumberOfRetryQueueItemsProcessed)]: %1.
#
# Events logged only to crimson (windows) event log
#
    TSDetectionStarted=The troubleshooter started detection.
    TSDetectionFinished=The troubleshooter finished detection.
    TSDetectionFailed=The troubleshooter failed during detection.
    
    TSResolutionStarted=The troubleshooter started resolution.
    TSResolutionFinished=The troubleshooter finished resolution.
    TSResolutionFailed=The troubleshooter failed during resolution. Reason: %1
###PSLOC
'@