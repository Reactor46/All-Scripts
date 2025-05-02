@ECHO ON
PAUSE

xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OMP\*.*" "\\msoit03\c$\reports\EDI\OMP" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\PAC\*.*" "\\msoit03\c$\reports\EDI\PAC" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\USON\*.*" "\\msoit03\c$\reports\EDI\USON" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ONC\*.*" "\\msoit03\c$\reports\EDI\ONC" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ORTHO\*.*" "\\msoit03\c$\reports\EDI\ORTHO" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OCMG PULM\*.*" "\\msoit03\c$\reports\EDI\PULM" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\RAD\*.*" "\\msoit03\c$\reports\EDI\RAD" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\USON\*.*" "\\10.20.3.33\Company\BILLING DEPARTMENT\BBP_Files" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OMP\*.*" "\\10.20.3.33\Company\~OMP\BBP Files" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\PAC\*.*" "\\10.20.3.33\Company\^PAC.AAC\BBP FILES" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ORTHO\*.*" "\\10.20.3.33\Company\~University Orthopaedics & Spine\BBP Files" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OCMG PULM\*.*" "\\10.20.3.33\Company\~Pulmonology\BBP Files" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ONC\*.*" "\\10.20.3.33\Company\~Nevada Cancer Specialists\BBP Files" >> C:\Scripts\TEST_LOG.LOG
PAUSE
xcopy /V/Z/Y "\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\RAD\*.*" "\\10.20.3.33\Company\~ROCN\BBP Files" >> C:\Scripts\TEST_LOG.LOG
PAUSE
EXIT


