ConvertFrom-StringData @'
###PSLOC
# Event Log strings
    DatabaseSpaceTroubleShooterStarted=The database space troubleshooter started on volume %1 for database %2.
    DatabaseSpaceTroubleShooterFoundLowSpace=The database space troubleshooter finished on volume %1 for database %2. The database is over the expected threshold. Users were quarantined to avoid running out of space. \nEDB free space (drive space + free space): %3 GB \nPercent EDB free space (drive space + free space): %4% \nLog drive free space: %5 GB \nPercent log drive free space: %6% \nEDB free space threshold: %7% \nLog free space threshold: %8% \nHour threshold: %9 Hrs \nGrowth rate threshold: %10 B/Hr \nInitial growth rate: %11 B/Hr \nFinal growth rate: %12 B/Hr \nNumber of users quarantined: %13 \nPercent EdbFreeSpaceCriticalThreshold: %14% \nPercent EdbFreeSpaceAlertThreshold: %15% \nQuarantine: %16
    DatabaseSpaceTroubleShooterFoundLowSpaceNoQuarantine=The database space troubleshooter finished on volume %1 for database %2. The database is over the expected threshold, but is not growing at an unusual rate. No quarantine action was taken. \nEDB free space (drive space + free space): %3 GB \nPercent EDB free space (drive space + free space): %4% \nLog drive free space: %5 GB \nPercent log drive free space: %6% \nEDB free space threshold: %7% \nLog free space threshold: %8% \nHour threshold: %9 Hrs \nGrowth rate threshold: %10 B/Hr \nInitial growth rate: %11 B/Hr \nFinal growth rate: %12 B/Hr \nPercent EdbFreeSpaceCriticalThreshold: %13% \nPercent EdbFreeSpaceAlertThreshold: %14% \nQuarantine: %15
    DatabaseSpaceTroubleShooterFinishedInsufficient=The database space troubleshooter finished on volume %1 for database %2. The database is over the expected threshold and continues to grow. Manual intervention is required. \nEDB free space (drive space + free space): %3 GB \nPercent EDB free space (drive space + free space): %4% \nLog drive free space: %5 GB \nPercent log drive free space: %6% \nEDB free space threshold: %7% \nLog free space threshold: %8% \nHour threshold: %9 Hrs \nGrowth rate threshold: %10 B/Hr \nInitial growth rate: %11 B/Hr \nFinal growth rate: %12 B/Hr \nNumber of users quarantined: %13 \nPercent EdbFreeSpaceCriticalThreshold: %14% \nPercent EdbFreeSpaceAlertThreshold: %15% \nQuarantine: %16
    DatabaseSpaceTroubleShooterNoProblemDetected=The database space troubleshooter finished on volume %1 for database %2. No problems were detected. \nEDB free space (drive space + free space): %3 GB \nPercent EDB free space (drive space + free space): %4% \nLog drive free space: %5 GB \nPercent log drive free space: %6% \nEDB free space threshold: %7% \nLog free space threshold: %8% \nHour threshold: %9 hrs \nCurrent growth rate: %10 B/hr \nPercent EdbFreeSpaceCriticalThreshold: %11% \nPercent EdbFreeSpaceAlertThreshold: %12% \nQuarantine: %13
    DatabaseSpaceTroubleShooterQuarantineUser=The database space troubleshooter quarantined mailbox %1 in database %2.
    DatabaseSpaceTroubleDetectedAlertSpaceIssue=The database space troubleshooter detected a low space condition on volume %1 for database %2. Provisioning for this database has been disabled. Database is under %3% free space.
    DatabaseSpaceTroubleDetectedCriticalSpaceIssue=The database space troubleshooter has detected a critically low space condition on volume %1 for database %2. Provisioning for this database has been disabled. The database has less than %3% free space.
    DatabaseSpaceTroubleDetectedWarningSpaceIssue=The database space troubleshooter detected a low space condition on volume %1 for database %2. Provisioning for this database has been disabled. Database is under %3% free space.

    DatabaseLatencyTroubleShooterStarted=The database latency troubleshooter started on database %1.
    DatabaseLatencyTroubleShooterNoLatency=The database latency troubleshooter detected that the current latency of %1 ms for database %2 is within the threshold of %3 ms.
    DatabaseLatencyTroubleShooterLowOps=The database latency troubleshooter detected that the current latency for database %1 appears high, but the current load on the database is too low for this metric to be meaningful. \n\nCurrent latency: %2 ms. (Usual threshold is %3 ms.) \nCurrent load: %4 operations per second. (Minimum load for meaningful latency: %5 ops/sec.)
    DatabaseLatencyTroubleShooterBadDiskLatencies=The database latency troubleshooter detected that disk latencies are abnormal for database %1. You need to replace the disk. \nRead Latency: %2 \nRead Rate: %3 \nRPC Average Latency: %4
    DatabaseLatencyTroubleShooterBadDSAccessActiveCallCount=The database latency troubleshooter detected that the DSAccess Active Call Count is abnormal for database %1. This may be due to an Active Directory problem. \nDSAccess Average Latency: %2 \nDSAccess Active Calls: %3 \nRPC Average Latency: %4
    DatabaseLatencyTroubleShooterBadDSAccessAverageLatency=The database latency troubleshooter detected that the DSAccess Average Latency is abnormal for database %1. This may be due to an Active Directory problem. \nDSAccess Average Latency: %2 \nDSAccess Active Calls: %3 \nRPC Average Latency: %4
    DatabaseLatencyTroubleShooterQuarantineUser=The database latency troubleshooter quarantined user %1 on database %2 due to unusual activity in the mailbox. If the problem persists, manual intervention will be required. \nAverage time in server: %3 \nRPC Average Latency: %4 \nRPC Operations per second: %5
    DatabaseLatencyTroubleShooterNoQuarantine=The database latency troubleshooter identified a problem with user %1 on database %2 due to unusual activity in the mailbox. No quarantine has been performed since the Quarantine parameter wasn't specified. If the problem persists, manual intervention is required. \nAverage time in server: %3 \nRPC Average Latency: %4 \nRPC Operations per second: %5
    DatabaseLatencyTroubleShooterIneffective=The database latency troubleshooter detected high RPC Average latencies for database %1 but was unable to determine a cause. Manual intervention is required. \nRPC Average Latency: %2 \nRPC Operations per second: %3
###PSLOC
'@