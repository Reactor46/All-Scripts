
ECHO Memory Usage For lassvc01
ECHO
TaskList /FI "IMAGENAME eq creditone*" /S lassvc01
TaskList /FI "IMAGENAME eq creditone*" /S lassvc01 /SVC
TaskList /FI "IMAGENAME eq ASPNET*" /S lassvc01
Pause

