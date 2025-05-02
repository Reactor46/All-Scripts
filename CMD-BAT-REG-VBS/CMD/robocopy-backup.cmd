@ECHO ON
SETLOCAL

SET _source=E:\home

SET _dest=\\voltron\sbe\RSYNCH\RSYNC-BACKUP\acu1\home

SET _what=/COPYALL /B /S /E /MIR
:: /COPYALL :: COPY ALL file info
:: /B :: copy files in Backup mode. 
:: /SEC :: copy files with SECurity
:: /MIR :: MIRror a directory tree 

SET _options=/ZB /R:2 /W:2 /LOG:acu1_home.txt /TEE
:: /R:n :: number of Retries
:: /W:n :: Wait time between retries
:: /LOG :: Output log file
:: /NFL :: No file logging
:: /NDL :: No dir logging

ROBOCOPY %_source% %_dest% %_what% %_options%