FULL BACK UP

@echo off
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)
@echo *************** Gavans Backup Script{Exchange} ***************
@echo *
@echo *   C:\Company Data
@echo *   Maximum compression using z7ip
@echo *   Archive saved to g:\%mydate%-company_data_FULL.7z
@echo *
@echo **************************************************************
echo [%time:~0,5%] - Beginning full Backup. >> g:\%mydate%"-BackupLog_FULL.txt
c:\hobocopy /verbosity=3 /full /skipdenied /y /r /statefile=g:\Company_Data\Company_Data.dat c:\Juncture_Point_DO_NOT_DELETE_SEE_GAVAN g:\%mydate%
echo [%time:~0,5%] - Finished full Backup. >> g:\%mydate%"-BackupLog_FULL.txt
echo [%time:~0,5%] - Begining Compressing full backup to %mydate%-company_data_FULL.7z >> g:\%mydate%"-BackupLog_FULL.txt
7z a -mx9 g:\%mydate%-company_data_FULL.7z g:\%mydate% -pSECRET
echo [%time:~0,5%] - Finished Compressing full backup to %mydate%-company_data_FULL.7z >> g:\%mydate%"-BackupLog_FULL.txt
echo [%time:~0,5%] - Deleted g:\%mydate% >> g:\%mydate%"-BackupLog_FULL.txt
echo [%time:~0,5%] - Backup Successfully Finished >> g:\%mydate%"-BackupLog_FULL.txt

INCREMENTAL BACKUP

@echo off
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)
@echo *************** Gavans Backup Script{Exchange} ***************
@echo *
@echo *   C:\Company Data
@echo *   Maximum compression using z7ip
@echo *   Archive saved to g:\%mydate%-company_data.7z
@echo *
@echo **************************************************************
echo [%time:~0,5%] - Beginning Incremental Backup. >> g:\%mydate%"-BackupLog.txt
c:\hobocopy /verbosity=3 /incremental /skipdenied /statefile=g:\Company_Data\Company_Data.dat /y /r c:\Juncture_Point_DO_NOT_DELETE_SEE_GAVAN g:\%mydate%
echo [%time:~0,5%] - Finished Incremental Backup. >> g:\%mydate%"-BackupLog.txt
echo [%time:~0,5%] - Deleting empty directorys >> g:\%mydate%"-BackupLog.txt
for /f "usebackq delims=" %%d in (`"dir /ad/b/s/q g:\%mydate% | sort /R"`) do rd "%%d"
echo [%time:~0,5%] - Finished Deleting empty directorys >> g:\%mydate%"-BackupLog.txt
echo [%time:~0,5%] - Begining Compressing incremental backup to %mydate%-company_data.7z >> g:\%mydate%"-BackupLog.txt
7z a -mx9 g:\%mydate%-company_data.7z g:\%mydate% -pSECRET
echo [%time:~0,5%] - Finished Compressing incremental backup to %mydate%-company_data.7z >> g:\%mydate%"-BackupLog.txt
move g:\%mydate%-company_data.7z h:\
echo [%time:~0,5%] - Moved %mydate%-company_data.7z to H:\ >> g:\%mydate%"-BackupLog.txt
echo _________________________________BACKUP_BEGIN_INDEX_________________________________ >> g:\%mydate%"-BackedUpFileIndex.txt
dir /s "g:\%mydate%" >> g:\%mydate%"-BackedUpFileIndex.txt 
echo _________________________________BACKUP_END_INDEX_________________________________ >> g:\%mydate%"-BackedUpFileIndex.txt
echo [%time:~0,5%] - Please see g:\%mydate%"-BackedUpFileIndex.txt for an index of files backed up.rmdir /s /q "g:\%mydate%"
echo [%time:~0,5%] - Deleted g:\%mydate% >> g:\%mydate%"-BackupLog.txt
echo [%time:~0,5%] - Backup Successfully Finished >> g:\%mydate%"-BackupLog.txt