@ECHO ON
SETLOCAL

FULL BACK UP

@echo off
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)
@echo *************** Johns Robocopy Script ***************
echo [%time:~0,5%] - Beginning full sync. >> C:\Scripts\Logs\Robocopy.%mydate%.log

SET _source1="\\server\C$\Users Shared Folders"
SET _dest1="C:\Users Shared Folders"

SET _source2="\\server\C$\safedb"
SET _dest2="C:\safedb"

SET _source3="\\server\C$\ned2000"
SET _dest3="C:\ned2000"

SET _source4="\\server\C$\medemr"
SET _dest4="C:\medemr"

SET _source5="\\server\C$\MED2000Local"         
SET _dest5="C:\Med2000Local"

SET _source6="\\server\C$\Med2000"              
SET _dest6="C:\Med2000"

SET _source7="\\server\C$\kyocera"              
SET _dest7="C:\Kyocera"

SET _source8="\\server\C$\etemp"                
SET _dest8="C:\etemp"

SET _source9="\\server\C$\emrsafe"              
SET _dest9="C:\emrsafe"

SET _source10="\\server\C$\emailtest"            
SET _dest10="C:\emailtest"

SET _source11="\\server\C$\Elevate Software"     
SET _dest11="C:\Elevate Software"

SET _source12="\\server\C$\EDBServer"            
SET _dest12="C:\EDBServer"

SET _source13="\\server\C$\eback"               
SET _dest13="C:\eback"

SET _source15="\\server\C$\cs_test"
SET _dest15="C:\cs_test"

SET _source16="\\server\C$\ClientApps"
SET _dest16="C:\ClientApps"

SET _source17="\\server\C$\aug13"                
SET _dest17="C:\aug13"

SET _source18="\\server\C$\anna"                 
SET _dest18="C:\anna"

SET _source19="\\server\C$\admin"
SET _dest19="C:\admin"

SET _source20="\\server\C$\Accounting"           
SET _dest20="C:\Accounting"

SET _source21="\\server\C$\aaa"
SET _dest21="C:\aaa"



SET _what=/COPYALL /B /S /E /MIR
:: /COPYALL :: COPY ALL file info
:: /B :: copy files in Backup mode. 
:: /SEC :: copy files with SECurity
:: /MIR :: MIRror a directory tree 

SET _options=/ZB /R:2 /W:2
:: /R:n :: number of Retries
:: /W:n :: Wait time between retries
:: /LOG :: Output log file
:: /NFL :: No file logging
:: /NDL :: No dir logging

ROBOCOPY %_source1% %_dest1% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source2% %_dest2% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source3% %_dest3% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source4% %_dest4% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source5% %_dest5% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source6% %_dest6% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source7% %_dest7% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source8% %_dest8% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source9% %_dest9% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source10% %_dest10% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source11% %_dest11% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source12% %_dest12% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source13% %_dest13% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source14% %_dest14% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source15% %_dest15% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source16% %_dest16% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source17% %_dest17% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source18% %_dest18% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source19% %_dest19% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source20% %_dest20% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log
ROBOCOPY %_source21% %_dest21% %_what% %_options% >> C:\Scripts\Logs\Robocopy.%mydate%.log

echo [%time:~0,5%] - Finished full Backup. >> C:\Scripts\Logs\Robocopy.%mydate%.log