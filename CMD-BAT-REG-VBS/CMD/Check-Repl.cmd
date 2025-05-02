# Run on USONVSVRDC01
repadmin /syncall /Apedq USONVSVRDC02 CN=Configuration,DC=USON,DC=LOCAL
repadmin /syncall /Apedq USONVSVRDC03 CN=Configuration,DC=USON,DC=LOCAL
# Run on USONVSVRDC02
repadmin /syncall /Apedq USONVSVRDC01 CN=Configuration,DC=USON,DC=LOCAL
repadmin /syncall /Apedq USONVSVRDC03 CN=Configuration,DC=USON,DC=LOCAL
# Run on USONVSVRDC03
repadmin /syncall /Apedq USONVSVRDC01 CN=Configuration,DC=USON,DC=LOCAL
repadmin /syncall /Apedq USONVSVRDC02 CN=Configuration,DC=USON,DC=LOCAL


repadmin /showrepl USONVSVRDC01
repadmin /showrepl USONVSVRDC02
repadmin /showrepl USONVSVRDC03

repadmin /replsummary *

repadmin /showutdvec USONVSVRDC0* DC=USON,DC=LOCAL /latency



repadmin /showattr fsmo_dnm: ncobj:config: /subtree /filter:(objectClass=crossRef) /atts:nCName,dnsRoot,net,dnsRoot,net,biosname,systemFlags /homeserver:USONVSVRDC01
repadmin /showattr fsmo_dnm: ncobj:config: /subtree /filter:(objectClass=crossRef) /atts:nCName,dnsRoot,net,dnsRoot,net,biosname,systemFlags /homeserver:USONVSVRDC02
repadmin /showattr fsmo_dnm: ncobj:config: /subtree /filter:(objectClass=crossRef) /atts:nCName,dnsRoot,net,dnsRoot,net,biosname,systemFlags /homeserver:USONVSVRDC03