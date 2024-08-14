
ECHO Memory Usage For lascapsmt01
ECHO
TaskList /FI "IMAGENAME eq FNBM*" /S lascapsmt01
TaskList /FI "IMAGENAME eq FNBM*" /S lascapsmt01 /SVC
TaskList /FI "IMAGENAME eq ASPNET*" /S lascapsmt01
Pause

