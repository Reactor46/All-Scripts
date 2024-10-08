# Localized	09/03/2016 06:56 AM (GMT)	303:4.80.0411 	CITSLibrary.strings.psd1
ConvertFrom-StringData @'
###PSLOC
    # English strings
AllNotAllowedForResolve=When 'Resolve' is specified for the Action parameter, 'All' is not allowed for the Symptom parameter.
TroubleshooterFailed=The troubleshooter failed with error:
RegistryOpenError=Couldn't open the registry on server:
RegistryReadError=Couldn't read registry key:
TimeoutWaitingForEvent=Event couldn't be found within the timeout period:
TimeoutWaitingForProcessToStop=The process didn't stop within the time-out period:
TimeoutZeroOrNegative=Timeout reached zero or negative value entering function:

# Event Log strings 
# Logged both to application and crimson logs   
#
TSStarted=The Troubleshooter started successfully.
DetectedDeadlock=Detected search service deadlock.
DetectedCatalogCorruption=Detected catalog corruption for database %1
DetectedIndexingStall=Detected indexing stall for database %1
DetectedIndexingBacklog=Indexing backlog reached a critical limit of %2 hours or more for database %1
TSFailed=The troubleshooter failed with exception %1.
TSSuccess=The troubleshooter finished successfully.
RestartSuccess=Restart of search services succeeded
RestartFailure=Search services failed to restart. Reason: %1
DetectedNoIssues=The troubleshooter didn't find any issues for any catalog.
CatalogHasNoIssues=The troubleshooter didn't find any catalog issues for database %1.
DetectedSameSymptomTooManyTimes=The troubleshooter detected the symptom %1 %2 times in the past %3 hours for catalog %4. This exceeded the allowed limit for failures.
ReseedSuccess=Reseeding succeeded for the catalog of database %1.
ReseedFailure=Seeding of the content index catalog for mailbox database %1 failed with the following reason: %2.
AnotherInstanceRunning=Another instance of the troubleshooter is already running on this computer. Two or more instances cannot be run simultaneously.
#
# Events logged only to crimson (windows) event log
#
TSDetectionStarted=The troubleshooter started detection.
TSDetectionFinished=The troubleshooter finished detection.
TSDetectionFailed=The troubleshooter failed during detection.
    
TSResolutionStarted=The troubleshooter started resolution.
TSResolutionFinished=The troubleshooter finished resolution.
TSResolutionFailed=The troubleshooter failed during reosolution. Reason: %1
###PSLOC
'@

