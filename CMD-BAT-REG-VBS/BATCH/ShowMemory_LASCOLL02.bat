
ECHO Memory Usage For lascoll02
ECHO
Tasklist /FI "IMAGENAME eq Contoso*" /S lascoll02
Tasklist /FI "IMAGENAME eq Contoso*" /S lascoll02 /SVC
Tasklist /FI "IMAGENAME eq w3wp*" /S lascoll02
Tasklist /FI "IMAGENAME eq w3wp*" /S lascoll02 /SVC
Pause

