
ECHO Memory Usage For %1
ECHO
TaskList /FI "IMAGENAME eq Contoso*" /S %1
TaskList /FI "IMAGENAME eq Contoso*" /S %1 /SVC
TaskList /FI "IMAGENAME eq ASPNET*" /S %1
Pause

