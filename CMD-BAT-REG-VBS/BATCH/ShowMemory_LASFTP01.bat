
ECHO Memory Usage For lasftp01
ECHO
TaskList /FI "IMAGENAME eq cftp*" /S lasftp01
TaskList /FI "IMAGENAME eq cftp*" /S lasftp01 /SVC1
Pause