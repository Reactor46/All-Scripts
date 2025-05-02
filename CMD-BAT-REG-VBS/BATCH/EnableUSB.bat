@ECHO OFF
REM  QBFC Project Options Begin
REM  HasVersionInfo: No
REM  Companyname: 
REM  Productname: 
REM  Filedescription: 
REM  Copyrights: 
REM  Trademarks: 
REM  Originalname: 
REM  Comments: 
REM  Productversion:  0. 0. 0. 0
REM  Fileversion:  0. 0. 0. 0
REM  Internalname: 
REM  Appicon: 
REM  AdministratorManifest: No
REM  QBFC Project Options End
ECHO ON
@echo off
echo .
echo ........Doing settings for USB Enable........
echo .
@pause

echo .
echo ........Registry entries........
echo .

REG ADD HKLM\SYSTEM\CURRENTCONTROLSET\SERVICES\USBSTOR /V Start /T REG_DWORD /D 3 /F

echo .
echo ........Group Policy edits........
echo .

REG ADD HKLM\Software\Policies\Microsoft\Windows\deviceInstall\Restrictions /V DenydeviceClassesRetroactive /T REG_DWORD /D 0 /F

echo .
echo ........USB Mass Storage has been enabled........
echo .

@pause
