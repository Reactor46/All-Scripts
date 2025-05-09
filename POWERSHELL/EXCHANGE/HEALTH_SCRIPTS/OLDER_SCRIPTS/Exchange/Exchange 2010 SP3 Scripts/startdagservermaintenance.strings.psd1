ConvertFrom-StringData @'
###PSLOC
res_0000 = Running 'cluster.exe node {0} /pause' to suspend the ability of the server to host the Primary Active Manager.
res_0001 = (Would normally run the Command now, but DagScriptTesting is set)
res_0002 = {1}: Could not contact the server {0} to suspend it from hosting the Primary ActiveManager. Ignoring and continuing. EPT_S_NOT_REGISTERED.
res_0003 = {1}: Could not contact the server {0} to suspend it from hosting the Primary ActiveManager. Ignoring and continuing. RPC_S_SERVER_UNAVAILABLE.
res_0004 = {2}: Failed to suspend the server {0} from hosting the Primary Active Manager, returned {1}.
res_0005 = {1}: Running 'GetCriticalMailboxResources {0}' the first time...
res_0006 = {1}: GetCriticalMailboxResources returned {0} objects.
res_0007 = The following objects are hosted by '{1}', before attempting to move them off: `n{0})
res_0008 = Setting DatabaseCopyAutoActivationPolicy to Blocked on server {0}
res_0009 = whatif: {1} {2} {0} {3}:Blocked
res_0010 = {1}: Suspending all passive database copies on server '{0}'...
res_0011 = Whatif: {4} `"{0}\\{2}`" {5} {6}:{3} {7} `"Suspended ActivationOnly by StartDagServerMaintenance.ps1 at {1}`"
res_0012 = {1} did not find any replicated databases on {0}.
res_0013 = {1}: Running 'GetCriticalMailboxResources {0} ' second time around to determine if some critical resources are still there...
res_0014 = The following objects are still hosted by '{1}', even after attempting to move them off: `n{0})
res_0015 = The script encountered some error! Attempting to roll back any changes performed.
res_0016 = Restoring auto activation policy on the server {0}
res_0017 = whatif: (Would normally run the {0} command now)
res_0018 = Running 'cluster.exe /cluster:{0} node {1} /resume' to resume the ability of the server {2} to host the Primary Active Manager.
res_0019 = (whatif: Would normally run the cluster.exe Command now)
res_0020 = The server {0} is already able to host the Primary Active Manager.
res_0021 = {1}: Could not contact the server {0} to resume it to be able to host the Primary ActiveManager. Ignoring and continuing. EPT_S_NOT_REGISTERED.
res_0022 = {1}: Could not contact the server {0} to resume it to be able to host the Primary ActiveManager. Ignoring and continuing. RPC_S_SERVER_UNAVAILABLE.
res_0023 = {4}: Failed to resume the ability of the server {0} to host the Primary Active Manager, 'cluster.exe /cluster:{1} node {2} /resume' returned {3}.
res_0024 = StopDagServerMaintenance: Resuming mailbox database copying on {0}\\{1}. This clears the Activation Suspended state.
###PSLOC
'@
