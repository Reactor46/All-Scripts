@echo off
   
   :: This script is a generic unattended installer/uninstaller. It helps you to
   :: run installer.cmd with the right command line arguments. In addition it
   :: allows you to simply add *-preinstall.cmd *-postinstall.cmd scripts:
   :: call trace:
   ::  - unattended.cmd                | - unattended-uninstall.cmd
   ::   -> unattended-preinstall.cmd   |   -> unattended-uninstall.cmd
   ::   -> installing application      |   -> remove application
   ::   -> unattended-postinstall.cmd  |   -> unattended-uninstall-postinstall.cmd
   
   
   :: Name of the application (just to print it on the command prompt)
   set PROGRAM_NAME=LibreOffice
   
   :: 32-bit installer command (run on 32-bit Windows)
   if not "%CMD32%" == "" goto CMD32alreadySet
   set CMD32=LibO_3.5.0_Win_x86_install_multi.msi
   :CMD32alreadySet
   
   :: 64-bit installer command (run on 64-bit Windows)
   if not "%CMD64%" == "" goto CMD64alreadySet
   :: set to %CMD32% to install the same package on 64-bit Windows
   set CMD64=%CMD32%
   :CMD64alreadySet
   
   :: Type of installer, select one supported by install.cmd
   :: e.g. msiinstall, msiuninstall, nsis, innosetup...
   set INSTALLER_TYPE=msiuninstall
   
   :: Additional options to be passed to installer.
   set INSTALLER_OPTIONS=
   
   :: Working directory for installer
   set INSTALLER_WORKDIR=
   
   :: install helper script name (needs to be within the same directory)
   set INSTALLER=install.cmd
   
   
   
   set PROGRAM_FILES=%ProgramFiles%
   if not "%ProgramFiles(x86)%" == "" set PROGRAM_FILES=%ProgramFiles(x86)%
   
   :: check if MS Office is installed
   if exist "%PROGRAM_FILES%\Microsoft Office" set FILEASSOC=0
   
   :: custom options to pass to the installer
   :: set CUSTOM_OPTIONS=SELECT_WORD=%FILEASSOC% SELECT_EXCEL=%FILEASSOC% SELECT_POWERPOINT=%FILEASSOC% ADDLOCAL=ALL,gm_o_Accessories REMOVE=gm_o_Quickstart,gm_o_Systemintegration,gm_o_Testtool,gm_o_Winexplorerext,gm_o_Winexplorerext_PropertyHdl
   :: set CUSTOM_OPTIONS=TRANSFORMS=%~dp0trans_de.mst SELECT_WORD=%FILEASSOC% SELECT_EXCEL=%FILEASSOC% SELECT_POWERPOINT=%FILEASSOC% ADDLOCAL=ALL REMOVE=gm_o_Quickstart,gm_o_Systemintegration,gm_o_Testtool,gm_o_Winexplorerext,gm_o_Winexplorerext_PropertyHdl
   
   
   :: ############################################################################
   :: No need to change anything below this line (usually ;-))
   :: ############################################################################
   set INSTALLER_LOC=%~dp0
   set CMDPATH=%~dpn0
   
   if exist "%INSTALLER_LOC%prerun.cmd" (
   	setlocal
   	call "%INSTALLER_LOC%prerun.cmd"
   	endlocal
   )
   
   if exist "%CMDPATH%-pre.cmd" (
   	setlocal
   	call "%CMDPATH%-pre.cmd"
   	endlocal
   
   )
   
   :install
   echo Installing %PROGRAM_NAME%
   
   set PROGRAM_FILES=%ProgramFiles%
   if not "%ProgramFiles(x86)%" == "" set PROGRAM_FILES=%ProgramFiles(x86)%
   
   call "%INSTALLER_LOC%%INSTALLER%" %INSTALLER_TYPE% "%CMD32%" "%CMD64%" "%INSTALLER_WORKDIR%" "%INSTALLER_OPTIONS%"
   set EXIT_CODE=%ERRORLEVEL%
   
   if exist "%CMDPATH%-post.cmd" (
   	setlocal
   	call "%CMDPATH%-post.cmd"
   	endlocal
   )
   
   if exist "%INSTALLER_LOC%postrun.cmd" (
   	setlocal
   	call "%INSTALLER_LOC%postrun.cmd"
   	endlocal
   )
   
   :end
   exit /B %EXIT_CODE%