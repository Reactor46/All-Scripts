
ECHO Memory Usage For %1
ECHO
TaskList /FI "IMAGENAME eq Cache*" /S lascasmt02
TaskList /FI "IMAGENAME eq datalayer*" /S lascasmt02
TaskList /FI "IMAGENAME eq Cache*" /S lascasmt02 /SVC
TaskList /FI "IMAGENAME eq datalayer*" /S lascasmt02 /SVC
Pause

