ConvertFrom-StringData @'
###PSLOC
res_0000 = {0} is installed on this machine.
res_0001 = {0} is not installed on this machine! Cannot run this script.
res_0002 = GetCriticalMailboxResources: Entering: `$serverName={0} serverIsMonitored={1}
res_0003 = {1}: Server '{0}' cluster service is down or not installed.
res_0004 = GetCriticalMailboxResources: Server '{1}' is part of DAG '{0}'
res_0005 = GetCriticalMailboxResources: Could not find a DAG that server {0} belongs to.
res_0006 = GetCriticalMailboxResources: There are {0} machines in the DAG. Skipping the Primary Active Manager check.
res_0007 = {0}: Checking for PAM role is skipped because cluster service is down.
res_0008 = GetCriticalMailboxResources: Checking for Primary Active Manager role by querying the cluster service with command: cluster.exe /cluster:{0} group `"cluster group`"
res_0009 = GetCriticalMailboxResources: cluster.exe did not succeed. Return value {0}. Output: {1}
res_0010 = GetCriticalMailboxResources: Could not contact the server {0} to tell if it can host the Primary ActiveManager. Because the server is not running, it obviously cannot host the Primary Active Manager. RPC_S_SERVER_UNAVAILABLE.
res_0011 = GetCriticalMailboxResources: Could not contact the server {0} to tell if it can host the Primary ActiveManager. Because the service is not running, it obviously cannot host the Primary Active Manager. EPT_S_NOT_REGISTERED.
res_0012 = GetCriticalMailboxResources: Found Primary Active Manager to be: {0}
res_0013 = GetCriticalMailboxResources: {0} is a critical resource: it is the Primary Active Manager.
res_0014 = GetCriticalMailboxResources: Server '{0}' is *NOT* part of any DAG, so we are skipping the Primary Active Manager check
res_0015 = GetCriticalMailboxResources: Checking for active databases, passive copies not suspended, and mailboxes for {1} databases hosted on server {0}...
res_0016 = GetCriticalMailboxResources: Found {0} databases on server {1}
res_0017 = Check Database Redundancy script {0} not found!
res_0018 = {4}: {3} is a replicated database. activeCopy is {0}, the status is {1} and activation suspended flag is {2}.
res_0019 = {4}: {3} is a critical resource. activeCopy is {0}, the status is {1} and activation suspended flag is {2}. (Critical resources are either active, or NOT in a suspended state.)
res_0020 = {2}: {1} has an active copy on {0}.
res_0021 = {5}: {3} is a critical resource. activeCopy is {0}, the status is {1} and activation suspended flag is {2}. Mailbox server {4} (Critical resources are either active, or NOT in a suspended state.)
res_0022 = {4}: {3} is a critical resource. activeCopy is {0}, the status is {1} and the activation suspended flag is {2}. (Critical resources are either active, or NOT in a suspended state.)
res_0023 = GetCriticalMailboxResources: {0}\\{1} is a critical resource, removing {2} copy hosted on {3} will critically affect redundancy of this database.
res_0024 = GetCriticalMailboxResources: {0} is NOT a replicated database.
res_0025 = GetCriticalMailboxResources: {0} is a critical resource: it contains {1} mailboxes.
res_0026 = GetCriticalMailboxResources: {0} is a critical resource: it contains {1} arbitration mailboxes.
res_0027 = GetCriticalMailboxResources: Leaving
res_0028 = {1}: There are {0} servers in the DAG.
res_0029 = UnmovedResource: {0}
res_0030 = UnmovedResource: mailbox {0} on db {1}
res_0031 = UnmovedResource: database {0} has status {1}
res_0032 = UnmovedResource: mailbox database {0}
res_0033 = {1}: Entering: `$Server={0}
res_0034 = {1}: Could not find a DAG that server {0} belongs to.
res_0035 = {1}: There are {0} machines in the DAG. Skipping the Primary Active Manager check.
res_0036 = {1}: Failed to get cluster group status. 'cluster.exe group' returned {0}.
res_0037 = {2}: 'cluster.exe {0} group' returned: `n{1}
res_0038 = {1}: Server {0} is hosting the Primary Active Manager, which will be moved.
res_0039 = {2}: Executing 'Cluster {0} group `"Cluster Group`" /MoveTo:{1}
res_0040 = whatif: skipping moving Primary Active Manager.
res_0041 = {2}: Server {1} is not hosting the Primary Active Manager, which is hosted by {0}
res_0042 = {0}: Leaving
res_0043 = {1}: Moving off all active replicated databases off server '{0}'
res_0044 = {1}: Moving the Primary Active Manager off server '{0}' if necessary...
res_0045 = {1}: Finished moving active databases and Primary Active Manager from server '{0}'
res_0046 = {1}: An error occurred while moving critical resources off server '{0}'
res_0047 = {2}: Entering: `$MailboxServer={0}, `$Database={1}
res_0048 = {2}: Active server for database '{1}' is '{0}'
res_0049 = {2}: moving database '{0}' off server '{1}'...
res_0050 = {2}: Database '{0}' failed to move off server '{1}'
res_0051 = {2}: The active copy of mailbox database '{0}' cannot be found on server {1}.
res_0052 = {1}: Moving all replicated active databases off server {0}
res_0053 = {2}: moving database '{0}' off server '{1}'
res_0054 = {2}: move of database '{0}' off server '{1}' ***FAILED***
res_0055 = {1}: There are no replicated active database copies on server {0}.
res_0056 = {0}: No valid mailbox server name or database name was provided in arguments. At least a mailbox server name is required.
res_0057 = {3}: Entering: `$db={0}, `$srcServer={1}, `$preferredTarget={2}
res_0058 = {3}: {1} copies out of {0} for database {2} will be attempted for move.
res_0059 = {3}: Executing '{4} {5} {1} {6} {0} {7} {8}:{2}'...
res_0060 = {3} whatif: {4} {5} {0} {6} {1}.MailboxServer {7} {8}:{2}
res_0061 = {3}: '{4} {5} {1} {6} {0} {7} {8}:{2}' did NOT return an error.
res_0062 = {3}: '{4} {5} {1} {6} {0} {7} {8}:{2}' returned an error.
res_0063 = {1}: No database copies for database '{0}' that meet the switchover criteria were found.
res_0064 = {1}: Database '{0}' *FAILED* to move! Now attempting to perform rollback to prevent a DB outage...
res_0065 = {1}: Database '{0}' *FAILED* to move!
res_0066 = {1}: Leaving (returning '{0}')
res_0067 = {3}: Testing move criteria for {0}, with `$Lossless={1} and `$CICheck={2} ...
res_0068 = {5}: Name='{0}', Status='{1}', CIStatus='{2}', CopyQueueLength={3}, ReplayQueueLength={4}
res_0069 = {2}: Verifying database '{0}', `$srcServer={1}
res_0070 = {3}: Mailbox database '{0}' has been moved successfully from server {1} to server {2} and is still mounted.
res_0071 = {2} whatif: Mailbox database '{0}' has not been moved from server {1}. It is still mounted there.
res_0072 = {2}: Mailbox database '{0}' has not been moved from server {1}. It is still mounted there.
res_0073 = {2}: Mailbox database '{0}' is successfully mounted on server {1}.
res_0074 = {2} whatif: The '{3}' operation for mailbox database '{0}' is complete, but the active copy is not mounted on server {1}.
res_0075 = {2}: The '{3}' operation for mailbox database '{0}' is complete, but the active copy is not mounted on server {1}.
res_0076 = {2} whatif: The '{3}' operation for mailbox database '{0}' is complete, but the active copy is now mounted on server {1} when it was originally dismounted.
res_0077 = {2}: The '{3}' operation for mailbox database '{0}' is complete, but the active copy is now mounted on server {1} when it was originally dismounted.
res_0078 = {3}: Mailbox database '{0}' has been moved successfully from server {1} to server {2} and is still dismounted.
res_0079 = {2} whatif: Mailbox database '{0}' has not been moved from server {1}. It is still dismounted there.
res_0080 = {2}: Mailbox database '{0}' has not been moved from server {1}. It is still dismounted there.
res_0081 = {2}: Mailbox database '{0}' is successfully dismounted on server {1}.
res_0082 = {1} whatif: Mailbox database object for '{0}' does not have the 'Mounted' field set. {2} requires a DB object retrieved via '{3} {4}'
res_0083 = {1}: Mailbox database object for '{0}' does not have the 'Mounted' field set. {2} requires a DB object retrieved via '{3} {4}'
res_0084 = {2}: Mounting DB '{0}' on server '{1}'
res_0085 = {1} whatif: {2} {3} {0}
res_0086 = {2}: DB '{0}' successfully mounted on server '{1}'
res_0087 = {2}: DB '{0}' *FAILED* to mount on server '{1}'
res_0088 = {4}: Skipping mount for DB '{0}' on server '{1}' since '{2}={3}'
res_0089 = {5}: Attempt 1: Moving DB '{0}' to server '{1}' with command: {6} {7} {2} {8} {3} {9}:{4}
res_0090 = {3} whatif: {4} {5} {0} {6} {1} {7}:{2}
res_0091 = {2}: Attempt 1: DB '{0}' successfully moved back to server '{1}'
res_0092 = {2}: Attempt 1: DB '{0}' *FAILED* to move back to server '{1}'
res_0093 = {5}: Attempt 2: Moving DB '{0}' to server '{1}' with command: {6} {7} {2} {8} {3} {9} {10}:{4}
res_0094 = {3} whatif: {4} {5} {0} {6} {1} {7} {8}:{2}
res_0095 = {2}: Attempt 2: DB '{0}' successfully moved back to server '{1}'
res_0096 = {2}: Attempt 2: DB '{0}' *FAILED* to move back to server '{1}'
res_0097 = {3} entered. dagName={0}, serverName={1}, clusterCommand={2}
res_0098 = Running '{0}' on {1}.
res_0099 = {2} inner block: cluster.exe returned {0}. Output is '{1}'.
res_0100 = {2}: cluster.exe returned {0}. Output is '{1}'.
res_0101 = {1}: Could not contact the server name {0}. RPC_S_SERVER_UNAVAILABLE (usually means the server is down).
res_0102 = {1}: Could not contact the server name {0}. EPT_S_NOT_REGISTERED (usually means the server is up, but clussvc is down).
res_0103 = {1}: Could not contact the server name {0}. RPC_S_SEC_PKG_ERROR (usually means the net name resource is down).
res_0104 = {1}: cluster.exe did not succeed, but {0} was not a {2} error code. Not attempting any other servers. This may be an expected error by the caller.
res_0105 = {1}: {2} was unable to find any DAGs named '{0}'!
res_0106 = {2}: {3} found {0} DAGs named '{1}'!
res_0107 = {3} dagName='{0}' serverName='{1}' is returning '{2}'.
res_0108 = Sleeping for {0} seconds...
res_0109 = Failed at command '{0}' with '{1}'
###PSLOC
'@
