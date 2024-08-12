@ECHO ON

PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\*.*" "\\10.20.3.33\Company\BILLING DEPARTMENT\BBP_Files" >> C:\Scripts\TEST_LOG.LOG

PAUSE
EXIT