SET mydate=%date:~6,4%-%date:~3,2%-%date:~0,2%
SET dailystuff=C:\Documents and Settings\Administrator\Desktop\Daily Stuff\
IF NOT EXIST "%dailystuff%%mydate%" MKDIR "%dailystuff%%mydate%" 