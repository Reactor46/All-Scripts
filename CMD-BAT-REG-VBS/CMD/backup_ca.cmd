################################################################################################
# FILENAME: backup_ca.cmd
# VERSION:  1.01
# DATE:     29.12.2012
# AUTHOR:   Fabian Müller
#
# THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
# OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
# FITNESS FOR A PARTICULAR PURPOSE.
#
# This sample is not supported under any Microsoft standard support program or service. 
# The script is provided AS IS without warranty of any kind. Microsoft further disclaims all
# implied warranties including, without limitation, any implied warranties of merchantability
# or of fitness for a particular purpose. The entire risk arising out of the use or performance
# of the sample and documentation remains with you. In no event shall Microsoft, its authors,
# or anyone else involved in the creation, production, or delivery of the script be liable for 
# any damages whatsoever (including, without limitation, damages for loss of business profits, 
# business interruption, loss of business information, or other pecuniary loss) arising out of 
# the use of or inability to use the sample or documentation, even if Microsoft has been advised 
# of the possibility of such damages.
################################################################################################

@ECHO OFF

SET backup=C:\certsrv_backup
SET install=install
SET DBbackup=database
SET REGbackup=registry
SET TEMPLATESbackup=templates
SET IIS7backup=IIS7

rd %backup% /S /Q

mkdir %backup%
mkdir %backup%\%DBbackup%
mkdir %backup%\%REGbackup%
mkdir %backup%\%TEMPLATESbackup%

Echo Backing up the Certification Authority Database
certutil -backupDB "%backup%\%DBbackup%"

Echo Backing up the registry keys
reg export HKLM\System\CurrentControlSet\Services\CertSvc %backup%\%REGBackup%\HKLM_certsvc.reg
reg export HKLM\System\CurrentControlSet\Services\CertSvc\Configuration %backup%\%REGBackup%\HKLM_certsvc_configuration.reg
certutil -getreg > %backup%\%REGBackup%\getreg.txt
certutil -v -getreg > %backup%\%REGBackup%\getreg_verbose.txt

Echo Backing up the CSPs of this Certification Authority
certutil –getreg CA\CSP > %backup%\%REGBackup%\HKLM_certsvc_CSP.txt

Echo Documenting all certificate templates published in AD
certutil -v -template > %backup%\%TEMPLATESbackup%\Templates.txt

Echo Documenting all certificate templates published at this CA
certutil -v -catemplates > %backup%\%TEMPLATESbackup%\CA_Templates.txt

Echo Documenting CA information
certutil -v -cainfo > %backup%\CAInfo.txt

Echo Backup of IIS7 configuration
Echo "Deleting old CA IIS 7 Backup..."
%windir%\system32\inetsrv\appcmd.exe delete backup "Test Backup"
Echo "Creating new CA IIS 7 Backup"
%windir%\system32\inetsrv\appcmd.exe add backup "Test Backup"
robocopy /MIR %windir%\system32\inetsrv\backup %backup%\%IIS7backup% /R:1 /W:1

Echo Backup of CAPolicy.inf
robocopy %windir% %backup%\%install% CAPolicy.inf /R:1 /W:1

Echo.
Echo ======================================================================
Echo BEAR IN MIND THAT YOU HAVE TO BACKUP THE PRIVATE KEYS OF THIS CA, TOO!
Echo You can back up the keys e.g. using "certutil.exe -backupKey".
Echo ======================================================================

PAUSE