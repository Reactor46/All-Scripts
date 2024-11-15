ConvertFrom-StringData @'
###PSLOC
res_0000 = {2}: Some active databases are not properly indexed or discovered. CountOfActiveDbs:{0}, NumberOfDbsDiscovered:{1}
res_0001 = {1}: Returning '{0}'.
res_0002 = {1}: Now searching for all the Mailbox Databases in DAG '{0}'...
res_0003 = {1}: Could not find mailbox server '{0}'!
res_0004 = {1}: Could not find Exchange server '{0}'!
res_0005 = {3}: Database '{1}' is active on server '{2}'. Mounted='{0}'
res_0006 = {2}: Database '{0}' is passive on server '{1}'
res_0007 = {2}: Server '{0}' is in AD site '{1}'
res_0008 = {1}: Running command: cluster.exe /cluster:{0} node
res_0009 = {0}: cluster.exe failed to contact the cluster. RPC_S_SERVER_UNAVAILABLE. The cluster network name resource may be offline or the cluster may have lost quorum.
res_0010 = {2}: cluster.exe did not succeed. Return value {1}. \nOutput: {0}
res_0011 = {1}: cluster.exe returned the following output:`n {0}
res_0012 = {0}: cluster.exe output returned no regex matches!
res_0013 = {2}: cluster.exe operation completed in {0} ms. Returning '{1}'.
res_0014 = {1}: Found DAG '{0}'.
res_0015 = {1}: Failed to query the cluster using the cluster netname of '{0}'! Querying DAG member servers instead.
res_0016 = {1}: Database availability group '{0}' has no members.
res_0017 = {2}: Enumerating DAG servers starting at index='{1}', server='{0}'.
res_0018 = {1}: Unable to communicate with database availability group '{0}'. Marking all members as 'Down'.
res_0019 = {1}: Overall operation completed in {0} ms.
res_0020 = {1}: Entering: `$serverName={0}
res_0021 = {1}: Leaving (returning '{0}')
res_0022 = {0}: Checking if the local server is the PAM...
res_0023 = {1}: Returning '{0}'.
res_0024 = {0}: No Active Directory sites were found.
res_0025 = Only 1 AD site '{0}' found. Balancing DBs by ActivationPreference instead...
res_0026 = Sites have started off balanced with a maximum difference in active databases of {0}.`n
res_0027 = Sites have started off *UNBALANCED* with a maximum difference in active databases of {0} !`n
res_0028 = Sites are still unbalanced! Now attempting to move some databases to less preferred copies. CurrentMaxDelta={0}, AllowedMaxDelta={1}
res_0029 = {1}: Starting iteration: {0}...
res_0030 = {1}: Descending sorted servers: {0}
res_0031 = {2}: MaxActives={0}, MinActives={1}
res_0032 = {0}: Moving databases to less preferred copies is allowed.
res_0033 = {1}: Considering databases on server {0}...
res_0034 = {2}: Following DBs can possibly be moved off {0}: {1}
res_0035 = {1}: Found no database to move off server {0} !
res_0036 = {1}: Sorted Servers: {0}
res_0037 = {2}: Database '{0}' does NOT have a copy on server '{1}' !
res_0038 = {1}: Database copy '{0}' has already been attempted for move. Skipping this copy.
res_0039 = {2}: Database copy '{0}' (AP={1}) is currently the active copy. Skipping this copy.
res_0040 = {4}: Database copy '{0}' (AP={1}) is NOT more preferred than active copy on '{2}' (AP={3})
res_0041 = {1}: Database '{0}' has not been moved.
res_0042 = Database '{0}' has had no move attempted EVEN THOUGH it is not active on its most preferred copy!
res_0043 = {2}: MaxActives({0}), MinActives({1}). Moving databases off of overloaded server first.
res_0044 = {2}: MaxActives({0}), MinActives({1}). Moving databases off in random order.
res_0045 = {0} `n
res_0046 = Database '{0}' FAILED to move!
res_0047 = Database '{0}' successfully moved from site '{1}' to site '{2}'.
res_0048 = {2}: Entering: `$mdb={0}, `$serverList={1}
res_0049 = {2}: Checking if database '{0}' is still active on the same server '{1}'...
res_0050 = {3}: Database '{0}' was not moved because it has apparently already been moved away from original active server '{1}' to '{2}'.
res_0051 = {1}: Leaving (return '{0}')
res_0052 = {1}: No servers were specified for database '{0}' to move to.
res_0053 = {2}: Database '{0}' is not being moved because the move target is the current active server: '{1}'
res_0054 = Considering move of '{0}' from '{1}' (AP = {2}) to '{3}' (AP = {4})...
res_0055 = Database '{2}' CANNOT be moved to server '{3}' because it would cause the AD sites to become MORE unbalanced. AllowedMaxSiteDelta={4}, CurMax={0}, CurMin={1}
res_0056 = Database '{2}' CANNOT be moved to server '{3}' because it would cause the AD sites to become unbalanced. AllowedMaxSiteDelta={4}, CurMax={0}, CurMin={1}
res_0057 = {2}: Attempting to move database '{0}' to server '{1}'...
res_0058 = {2}: Whatif mode: Hypothetically moved database '{0}' to server '{1}'!
res_0059 = {2}: Successfully moved database '{0}' to server '{1}'!
res_0060 = {4}: 'Move-ActiveMailboxDatabase -Identity '{0}' -ActivateOnServer {1} -Confirm:{2}' FAILED! Returned result: `n{3}
res_0061 = {2}: Mounting DB '{0}' on server '{1}'
res_0062 = {2}: DB '{0}' successfully mounted on server '{1}'
res_0063 = {2}: DB '{0}' *FAILED* to mount on server '{1}'
res_0064 = {4}: Skipping mount for DB '{0}' on server '{1}' since '{2}={3}'
res_0065 = {1}: Database '{0}' *FAILED* to move! Now attempting to perform rollback to prevent a DB outage...
res_0066 = {5}: Attempt 1: Moving DB '{0}' to server '{1}' with command: Move-ActiveMailboxDatabase -Identity {2} -ActivateOnServer {3} -Confirm:{4}
res_0067 = {2}: Attempt 1: DB '{0}' successfully moved back to server '{1}'
res_0068 = {2}: Attempt 1: DB '{0}' *FAILED* to move back to server '{1}'
res_0069 = {5}: Attempt 2: Moving DB '{0}' to server '{1}' with command: Move-ActiveMailboxDatabase -Identity {2} -ActivateOnServer {3} -SkipClientExperienceChecks -Confirm:{4}
res_0070 = {2}: Attempt 2: DB '{0}' successfully moved back to server '{1}'
res_0071 = {2}: Attempt 2: DB '{0}' *FAILED* to move back to server '{1}'
res_0072 = Server {0} is not participating in DB redistribution because it is activation blocked!
res_0073 = Server {0} is not participating in DB redistribution because it is *DOWN* according to clustering!
res_0074 = {1}: {0}
res_0075 = {1}: Target server '{0}' has an activation policy of 'Unrestricted'.
res_0076 = {2}: {1} completed in {0} ms
res_0077 = Compiling code...
res_0078 = Done!
res_0079 = Sleeping for {0} seconds...
res_0080 = Failed at command '{0}' with '{1}'
res_0081 = The local server '{0}' is not the Primary Active Manager! Skipping database balancing on this server.
res_0082 = {0}: A database availability group wasn't specified. Searching the local Mailbox server first...
res_0083 = {1}: Found local mailbox server '{0}'
res_0084 = {1}: Local server is part of DAG '{0}'
res_0085 = The local mailbox server '{0}' is NOT part of a DAG! Please specify the -DagName parameter.
res_0086 = The local server '{0}' is NOT a mailbox server. Please specify the -DagName parameter.
res_0087 = Please specify the -DagName parameter.
res_0088 = {1}: Searching for DAG '{0}'...
res_0089 = Could not find DAG matching '{0}'!
res_0090 = Please use one of the parameters to select what the script should do. Exiting.
res_0091 = Script run with -DotSourceMode. Exiting.
res_0092 = Starting: {0}
res_0093 = {0}: Database lookup
res_0094 = {0}: Server lookup
res_0095 = {0}: Database distribution calculation
res_0096 = {0}: Moving all databases to their preferred copies.
res_0097 = {0}: Shuffling all databases randomly
res_0098 = Target server '{0}' has an activation policy of '{1}'.
res_0099 = Target server '{0}' has an activation policy of '{1}', and it is not in the same site as source server '{2}'.
res_0100 = Moved as part of database redistribution ({0}).
res_0101 = Database '{0}' CANNOT be moved to server '{1}' because it failed validation checks! Error: {2}
res_0102 = Sites are now balanced! CurrentMaxDelta={0}, AllowedMaxDelta={1}
res_0103 = Databases are evenly balanced across non-activation-blocked servers in the DAG.
res_0104 = Database '{2}' successfully moved from '{0}' to '{1}'.
res_0105 = Target server '{0}' can't host any more active databases because the server reached the configured MaximumActiveDatabases limit of {1} database(s).
res_0106 = Target server '{0}' was reported by Windows Failover Clustering as Offline.
res_0107 = Database copy '{0}\{1}' hasn't been assigned an activation preference.
###PSLOC
'@
