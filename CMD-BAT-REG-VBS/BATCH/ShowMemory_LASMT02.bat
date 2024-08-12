
ECHO Memory Usage For lasmt02
ECHO
TaskList /FI "IMAGENAME eq Contoso*" /S lasmt02
TaskList /FI "IMAGENAME eq Contoso*" /S lasmt02 /SVC
TaskList /FI "IMAGENAME eq ASPNET*" /S lasmt02
Pause

