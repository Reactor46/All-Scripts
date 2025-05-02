@ECHO ON
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\OMP\*.*" "\\10.20.3.33\Company\~OMP\BBP Files" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\OCMG PULM\*.*" "\\10.20.3.33\Company\~Pulmonology\BBP Files" >> C:\Scripts\TEST_LOG.LOG
PAUSE
EXIT