
ECHO Memory Usage For %1
ECHO
TaskList /FI "IMAGENAME eq Cache*" /S lascasmt01
TaskList /FI "IMAGENAME eq datalayer*" /S lascasmt01
TaskList /FI "IMAGENAME eq Cache*" /S lascasmt01 /SVC
TaskList /FI "IMAGENAME eq datalayer*" /S lascasmt01 /SVC
Pause

