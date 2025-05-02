write-host "Testing MAPI Connectivity for ANEXMBX01" -foregroundcolor white
Test-MAPIConnectivity -Server ANEXMBX01
write-host "Testing MAPI Connectivity for ANEXMBX02" -foregroundcolor white
Test-MAPIConnectivity -Server ANEXMBX02
write-host "Testing MAPI Connectivity for ANEXMBX03" -foregroundcolor white
Test-MAPIConnectivity -Server ANEXMBX03
write-host "Testing OWA connectivity" -foregroundcolor white
Test-OwaConnectivity -URL:https://owa.dmacc.edu/owa -MailboxCredential:(get-credential DMACC\scmiles)
write-host "Testing mailflow between ANEXMBX01 and ANEXMBX02" -foregroundcolor white
Test-Mailflow ANEXMBX01 -TargetMailboxServer ANEXMBX02
write-host "Testing mailflow between ANEXMBX01 and ANEXMBX03" -foregroundcolor white
Test-Mailflow ANEXMBX01 -TargetMailboxServer ANEXMBX03
write-host "Testing mailflow between ANEXMBX02 and ANEXMBX03" -foregroundcolor white
Test-Mailflow ANEXMBX02 -TargetMailboxServer ANEXMBX03

