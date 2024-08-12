ECHO ON
SETLOCAL

SET _source=\\LASDAWH3\F$

SET _dest=\\LASDAWH4\F$

SET _what=/COPYALL /S /E /MIR /XX
:: /COPYALL :: COPY ALL file info
:: /B :: copy files in Backup mode. 
:: /SEC :: copy files with SECurity
:: /MIR :: MIRror a directory tree 

SET _options=/ZB /R:2 /W:2 /LOG+:LASDAWH_Copy.log /TS /FP /TEE /ETA
:: /R:n :: number of Retries
:: /W:n :: Wait time between retries
:: /LOG :: Output log file
:: /NFL :: No file logging
:: /NDL :: No dir logging

ROBOCOPY %_source% %_dest% %_what% %_options%
