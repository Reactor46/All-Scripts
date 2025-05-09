ConvertFrom-StringData @'
###PSLOC
res_0000 = Running 'cluster.exe /cluster:{0} node {1} /resume' to resume the ability of the server {2} to host the Primary Active Manager.
res_0001 = (whatif: Would normally run the cluster.exe Command now)
res_0002 = The server {0} is already able to host the Primary Active Manager.
res_0003 = {1}: Could not contact the server {0} to resume it to be able to host the Primary ActiveManager. Ignoring and continuing. EPT_S_NOT_REGISTERED.
res_0004 = {1}: Could not contact the server {0} to resume it to be able to host the Primary ActiveManager. Ignoring and continuing. RPC_S_SERVER_UNAVAILABLE.
res_0005 = StopDagServerMaintenance: Resuming mailbox database copying on {0}\\{1}. This clears the Activation Suspended state.
res_0006 = whatif: (Would normally run the {0} command now)
res_0007 = Restoring auto activation policy on the server {0}
res_0008 = {4}: Failed to resume the ability of the server {0} to host the Primary Active Manager, 'cluster.exe /cluster:{1} node {2} /resume' returned {3}.
###PSLOC
'@
