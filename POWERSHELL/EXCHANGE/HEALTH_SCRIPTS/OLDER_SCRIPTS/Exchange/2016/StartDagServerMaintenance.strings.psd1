# Localized	09/03/2016 06:55 AM (GMT)	303:4.80.0411 	StartDagServerMaintenance.Strings.psd1
ConvertFrom-StringData @'
###PSLOC
res_0000=Running 'cluster.exe node {0} /pause' to suspend the ability of the server to host the Primary Active Manager.
res_0001=(Would normally run the Command now, but DagScriptTesting is set)
res_0002={1}: Could not contact the server {0} to suspend it from hosting the Primary ActiveManager. Ignoring and continuing. EPT_S_NOT_REGISTERED.
res_0003={1}: Could not contact the server {0} to suspend it from hosting the Primary ActiveManager. Ignoring and continuing. RPC_S_SERVER_UNAVAILABLE.
res_0004={3}: Cluster node {0} /{1} failed. Error code {2} returned.
res_0005={1}: Running 'GetCriticalMailboxResources {0}' the first time...
res_0006={1}: GetCriticalMailboxResources returned {0} objects.
res_0007=The following objects are hosted by '{1}', before attempting to move them off: `n{0})
res_0008=Setting DatabaseCopyAutoActivationPolicy to Blocked on server {0}
res_0009=whatif: {1} {2} {0} {3}:Blocked
res_0010={1}: Suspending all passive database copies on server '{0}'...
res_0011=Whatif: {4} `"{0}\\{2}`" {5} {6}:{3} {7} `"Suspended ActivationOnly by StartDagServerMaintenance.ps1 at {1}`"
res_0012={1} did not find any replicated databases on {0}.
res_0013={1}: Running 'GetCriticalMailboxResources {0} ' second time around to determine if some critical resources are still there...
res_0014=The following objects are still critical on server '{1}', even after attempting to move them off: `n{0}). Errors from Move attempt: "{2}".If there are still active databases on the server please kill MSExchangeIS or Reboot the server to force the Actives off and retry Maintenance operation.
res_0015={0}: The script encountered some error and Force was not used! Attempting to roll back any changes performed.
res_0016=Restoring auto activation policy on the server {0}
res_0017=whatif: (Would normally run the {0} command now)
res_0018=Running 'cluster.exe /cluster:{0} node {1} /resume' to resume the ability of the server {2} to host the Primary Active Manager.
res_0019=(whatif: Would normally run the cluster.exe Command now)
res_0020=The server {0} is already able to host the Primary Active Manager.
res_0021={1}: Could not contact the server {0} to resume it to be able to host the Primary ActiveManager. Ignoring and continuing. EPT_S_NOT_REGISTERED.
res_0022={1}: Could not contact the server {0} to resume it to be able to host the Primary ActiveManager. Ignoring and continuing. RPC_S_SERVER_UNAVAILABLE.
res_0023={4}: Failed to resume the ability of the server {0} to host the Primary Active Manager, 'cluster.exe /cluster:{1} node {2} /resume' returned {3}.
res_0024=StopDagServerMaintenance: Resuming mailbox database copying on {0}\\{1}. This clears the Activation Suspended state.
res_0025={3}: Could not contact server {0} to {1} it in the cluster. Ignoring and continuing. {2}.
res_0026={2}: Cluster node {0} /{1} was successful.
res_0027={0}: The script encountered some errors and Force was used! No roll back of changes performed.
res_0028={0}: {1} encountered errors - {2}. Skipping {1} and moving on to the next maintenance steps.
res_0029=Setting ServerComponent - HighAvailability state InActive on server {0}
res_0030=whatif: {1} {2} {0} {3}:Blocked
res_0031=Restoring ServerComponent - HighAvailability state on server {0}
res_0032=whatif: (Would normally run the {0} command now)
res_0033=Set-ServerComponentState HighAvailability to InActive JobState - {0}
res_0034=Get-ServerComponentState did not return component: {0} for server: {1}. Skip setting component state.
res_0035=MailboxDatabase objects were unable to move off of this server because of blackout hours.
###PSLOC
'@
