@ECHO OFF
:: =================== CONFIG ============================================
:: FROM Path (can be a share or fully qualified dir)
SET fromdir=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OMP
SET fromdir1=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\PAC 
SET fromdir2=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\USON
SET fromdir3=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ONC
SET fromdir4=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ORTHO
SET fromdir5=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OCMG PULM
SET fromdir6=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\RAD
SET fromdir7=\\172.20.0.231\NextGenRoot\Prod\BBP_Files
SET fromdir8=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\OMP
SET fromdir9=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\PAC_AAC
SET fromdir10=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\ONC
SET fromdir11=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\ORTHO
SET fromdir12=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\OCMG PULM
SET fromdir13=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\RAD
SET fromdir14=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\USON
SET fromdir15=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OMP
SET fromdir16=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\PAC
SET fromdir17=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ORTHO
SET fromdir18=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OCMG PULM
SET fromdir19=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ONC
SET fromdir20=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\RAD

:: TO Path (can be a share or fully qualified dir)
SET todir=\\msoit03\reports\EDI\OMP
SET todir1=\\msoit03\reports\EDI\PAC
SET todir2=\\msoit03\reports\EDI\USON
SET todir3=\\msoit03\reports\EDI\ONC
SET todir4=\\msoit03\reports\EDI\ORTHO
SET todir5=\\msoit03\reports\EDI\PULM
SET todir6=\\msoit03\reports\EDI\RAD
SET todir7=\\10.20.3.33\Company\BILLING DEPARTMENT\BBP_Files
SET todir8=\\10.20.3.33\Company\~OMP\BBP Files
SET todir9=\\10.20.3.33\Company\^PAC.AAC\BBP FILES
SET todir10=\\10.20.3.33\Company\~Nevada Cancer Specialists\BBP Files
SET todir11=\\10.20.3.33\Company\~University Orthopaedics & Spine\BBP Files
SET todir12=\\10.20.3.33\Company\~Pulmonology\BBP Files
SET todir13=\\10.20.3.33\Company\~ROCN\BBP Files
SET todir14=\\10.20.3.33\Company\BILLING DEPARTMENT\BBP_Files
SET todir15=\\10.20.3.33\Company\~OMP\BBP Files
SET todir16=\\10.20.3.33\Company\^PAC.AAC\BBP FILES
SET todir17=\\10.20.3.33\Company\~University Orthopaedics & Spine\BBP Files
SET todir18=\\10.20.3.33\Company\~Pulmonology\BBP Files
SET todir19=\\10.20.3.33\Company\~Nevada Cancer Specialists\BBP Files
SET todir20=\\10.20.3.33\Company\~ROCN\BBP Files

:: Path for log file !!!USE TRAILING SLASH!!! (default is a dir called "logs" in this script's dir)
SET logdir=C:\Scripts\LOGS\
:: =================== CONFIG ============================================

:: =================== SCRIPT ============================================
:: !!! NO NEED TO EDIT BEYOND THIS POINT !!!
:: Make a date_time stamp like 20130125_134200
SET hh=%time:~0,2%
:: Add a zero when this is run before 10 am.
IF "%time:~0,1%"==" " set hh=0%hh:~1,1%
SET yyyymmdd_hhmmss=%DATE:~-4%%DATE:~4,2%%DATE:~7,2%_%hh%%time:~3,2%%time:~6,2%
:: Make a name for the log file
SET logfile=%logdir%%yyyymmdd_hhmmss%_%~n0.log
:: (If not exist) create a logdir directory
IF NOT EXIST %logdir% MKDIR %logdir%


	:: The Robocopy magic
	robocopy "%fromdir%" "%todir%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir1%" "%todir1%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir2%" "%todir2%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir3%" "%todir3%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir4%" "%todir4%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir5%" "%todir5%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir6%" "%todir6%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir7%" "%todir7%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir8%" "%todir8%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir9%" "%todir9%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir10%" "%todir10%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir11%" "%todir11%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir12%" "%todir12%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir13%" "%todir13%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir14%" "%todir14%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir15%" "%todir15%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir16%" "%todir16%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir17%" "%todir17%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir18%" "%todir18%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir19%" "%todir19%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
	robocopy "%fromdir20%" "%todir20%" /COPYALL /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
		
	:: =================== SCRIPT ============================================
EXIT