# Localized	01/08/2013 09:51 AM (GMT)	303:4.80.0411 	CheckDatabaseRedundancy.Strings.psd1
ConvertFrom-StringData @'
###PSLOC
res_0000={1}: Entering: `$serverName={0}
res_0001={1}: Leaving (returning '{0}')
res_0002={2}: '{0}' '{1}': Entering...
res_0003={4}: '{0}' '{1}': Copy '{2}\\{3}' not found! Assuming it has replay service down.
res_0004={4}: '{0}' '{1}': Copy '{2}\\{3}' is not active! Assuming active has replay service down.
res_0005={3}: '{1}' '{2}': Active copy '{0}' has replay service down.
res_0006={2}: '{0}' '{1}': Leaving, returning 'False'
res_0007={1}: Running command: cluster.exe /cluster:{0} node
res_0008={0}: cluster.exe failed to contact the cluster. RPC_S_SERVER_UNAVAILABLE. The cluster network name resource may be offline or the cluster may have lost quorum.
res_0009={2}: cluster.exe did not succeed. Return value {1}. \nOutput: {0}
res_0010={1}: cluster.exe returned the following output:\n {0}
res_0011={0}: cluster.exe output returned no regex matches!
res_0012={2}: cluster.exe operation completed in {0} ms. Returning '{1}'.
res_0013={1}: Found DAG '{0}'.
res_0014={1}: Failed to query the cluster using the cluster netname of '{0}'! Querying DAG member servers instead.
res_0015={1}: Database availability group '{0}' has no members.
res_0016={2}: Enumerating DAG servers starting at index='{1}', server='{0}'.
res_0017={1}: Unable to communicate with database availability group '{0}'. Marking all members as 'Down'.
res_0018={1}: Overall operation completed in {0} ms.
res_0019={1} returned only '{0}' servers.
res_0020={0}: Filtering out databases we are not going to check...
res_0021={1}: '{0}': Entering...
res_0022={1}: '{0}': Created empty database redundancy state entry.
res_0023={3}: '{2}': CurrentRedundancyCount={0}, LastRedundancyCount={1}
res_0024={2}: '{0}': Redundancy count is lower than specified threshold of '{1}'. Setting the state to 'Red'.
res_0025={2}: '{0}': time in green = {1} secs.
res_0026={2}: '{0}': time in red = {1} secs.
res_0027={1}: Reporting a Green event for database copy '{0}'
res_0028={1}: Reporting a RED event for database copy '{0}'!
res_0029={5}: Active copy '{4}' has Status:{0}, ErrorEventId:{1}, \nErrorMessage: {2}, \nSuspendComment: {3}
res_0030={1}: Database '{0}' is not replicated. It has only one configured database copy.
res_0031={1}: {0} Returning 'Failed'.
res_0032={1}: {0} Returning 'Warning'.
res_0033={8}: Passive copy '{6}' has Status:{0}, CopyQueueLength={1}, ReplayQueueLength={2}, DatabaseName={7}, ErrorEventId:{3}, \nErrorMessage: {4}, \nSuspendComment: {5}
res_0034={1}: returning '{0}' servers.
res_0035={1}( {0} ): Entering...
res_0036={2}( {1} ): operation completed in {0} ms.
res_0037={0}: Entering...
res_0038={1}: operation completed in {0} ms.
res_0039={1}: Started {0} async jobs.
res_0040={1}: Async operations completed in {0} ms.
res_0041=`$results = {0}
res_0042=Compiling code...
res_0043=Done!
res_0044=Sleeping for {0} seconds...
res_0045=Failed at command '{0}' with '{1}'
res_0046=Calling DataCenter {0} function...
res_0047=Calling {0}...
res_0048=Send mail failed!
res_0049=No Hub Server found!
res_0050=Calling {0}...
res_0051=Entering {2}: `$from={0}, `$pri={1}
res_0052={0} failed!
res_0053=Sending notification mail to: {0}
res_0054=Using SMTP client for '{0}', port={1}
res_0055=Sending mail from '{0}'...
res_0056=Mail sent!
res_0057=Exceeded {0} retries sending mail to {1}.
res_0058=Retrying to send mail to {0}.
res_0059=Send failed. Trying to use a different smtp client if possible...
res_0060=Skipping sending an email report since everything is healthy.
res_0061=Iteration {0} of the monitoring check completed successfully.
res_0062=Iteration {0} of the monitoring check FAILED due to an error.
res_0063=Starting iteration {0} of the monitoring check...
res_0064=Database '{0}' is a recovery database. Please specify a {1} database to check the health of.
res_0065=Could not find database matching '{0}'.
res_0066=Found mailbox server '{0}'.
res_0067=Filtering out databases matching the following regex: '{0}'
res_0068=Could not find server matching '{0}'.
res_0069=Found {0} databases...
res_0070=Skipping {0} since there are no databases to check!
res_0071=Iteration {0} of the monitoring check completed in {1} ms
res_0072={0}: {1} is required.
res_0073={0}: {1} cannot be empty.
res_0074={1}: {2} returned only '{0}' servers.
res_0075={3}: Active copy {0}\\{1} will be removed. Redundancy count should be at least {2} for active copy removal.
res_0076={2}: Passive copy {0}\\{1} will be removed.
res_0077={2}: Copy {0}\\{1} was not found.
res_0078=Script run with {0}. Exiting.
res_0079=Please specify a {0} address as well when {1} is used.
res_0080=Running once.
res_0081=Running many times...
res_0082=Starting: {0}
res_0083=Loading DataCenter script library '{0}'
res_0084=File '{0}\\{1}' is not present, so skipping sending a mail.
###PSLOC
'@
