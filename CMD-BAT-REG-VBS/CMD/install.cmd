@echo off
   
   :: Usage:
   :: install.cmd <type> <32-bit-installer> <64-bit-installer> [installer-location [custom-options]]
   :: where type is one of
   ::     msiinstall     Install the given MSI package
   ::     msiuninstall   Uninstall the given MSI package
   ::     msiupdate      Update the given MSP package
   ::     install4j      Install4J setup
   ::     innosetup      Inno setup
   ::     installshield  Install shield
   ::     nsis           Nullsoft install system (NSIS)
   ::     custom         Custom installer - options required in this case
   :: 32-bit-installer   Full file name (including extension) of 32-bit installer
   :: 64-bit-installer   Full file name (including extension) of 64-bit installer
   :: installer-location Path where the installers are stored, if empty assumes directory where install.cmd is
   :: custom-options     Replace the default installer options with the ones given
   
   :: do not export variables
   setlocal
   
   :: Additional options to be passed to the installer
   :: set CUSTOM_OPTIONS=
   
   :: Global variables
   set INSTALL_CMD=
   set EXIT_CODE=0
   
   :: Get command type
   set TYPE=%~1
   
   :: Get 32-bit installer name
   set CMD32=%~2
   
   :: Get 64-bit installer name
   set CMD64=%~3
   
   :: get file path
   set INSTALLER_PATH=%~dp0
   if not "%~4" == "" (
   set INSTALLER_PATH=%~4
   )
   
   set OPTIONS=
   if not "%~5" == "" (
   set OPTIONS=%~5
   )
   
   :: Detect which system is used
   if not "%ProgramFiles(x86)%" == "" goto 64bit
   goto 32bit
   
   
   :: ##########################################################################
   :: 64-bit system detected
   :: ##########################################################################
   :64bit
   :: Determine 64-bit installer to be used
   echo 64-bit system detected.
   :: set INSTALLER64=
   if not "%CMD64%" == "" (
   set INSTALLER64=%CMD64%
   ) else (
   :: Use 32-bit installer if available, no 64-bit installer available.
   if not "%CMD32%" == "" (
   echo Using 32-bit installer, no 64-bit installer specified.
   set INSTALLER64=%CMD32%
   ) else (
   echo Neither 64-bit nor 32-bit installer specified. Exiting.
   goto usage
   )
   )
   
   :: Check if installer is valid
   if exist "%INSTALLER_PATH%%INSTALLER64%" (
   set INSTALL_CMD=%INSTALLER_PATH%%INSTALLER64%
   ) else (
   if exist "%INSTALLER64%" (
   set INSTALL_CMD=%INSTALLER64%
   ) else (
   echo Installer "%INSTALLER_PATH%%INSTALLER64%" cannot be found! Exiting.
   exit /B 97
   )
   )
   goto installerselection
   
   
   :: ##########################################################################
   :: 32-bit system detected
   :: ##########################################################################
   :32bit
   :: Determine 32-bit installer to be used
   echo 32-bit system detected.
   set INSTALLER32=
   if not "%CMD32%" == "" (
   set INSTALLER32=%CMD32%
   ) else (
   echo No 32-bit installer specified. Exiting.
   exit /B 96
   )
   
   
   :: Check if installer is valid
   if exist "%INSTALLER_PATH%%INSTALLER32%" (
   set INSTALL_CMD=%INSTALLER_PATH%%INSTALLER32%
   ) else (
   if exist "%INSTALLER32%" (
   set INSTALL_CMD=%INSTALLER32%
   ) else (
   echo Installer "%INSTALLER_PATH%%INSTALLER32%" cannot be found! Exiting.
   exit /B 95
   )
   )
   goto installerselection
   
   
   
   :: ##########################################################################
   :: select installer system
   :: ##########################################################################
   :installerselection
   if /i "%TYPE%" == "msiinstall"    goto msiinstaller
   if /i "%TYPE%" == "msiinstall"    goto msiupdate
   if /i "%TYPE%" == "msiuninstall"  goto msiuninstaller
   if /i "%TYPE%" == "install4j"     goto install4j
   if /i "%TYPE%" == "innosetup"     goto innoinstaller
   if /i "%TYPE%" == "installshield" goto installshieldinstaller
   if /i "%TYPE%" == "nsis"          goto nsisinstaller
   if /i "%TYPE%" == "custom"        goto custominstaller
   goto usage
   
   
   
   :msiinstaller
   echo Installing "%INSTALL_CMD%"
   if "%OPTIONS%" == "" (
   set OPTIONS=/qn /norestart
   )
   start /wait "Software installation" /D"%INSTALLER_PATH%" msiexec /i "%INSTALL_CMD%" %OPTIONS% %CUSTOM_OPTIONS%
   set EXIT_CODE=%ERRORLEVEL%
   goto end
   
   :msiupdate
   echo Installing "%INSTALL_CMD%"
   if "%OPTIONS%" == "" (
   set OPTIONS=/qn /norestart
   )
   start /wait "Software installation" /D"%INSTALLER_PATH%" msiexec /update "%INSTALL_CMD%" %OPTIONS% %CUSTOM_OPTIONS%
   set EXIT_CODE=%ERRORLEVEL%
   goto end
   
   :msiuninstaller
   echo Uninstalling "%INSTALL_CMD%"
   if "%OPTIONS%" == "" (
   set OPTIONS=/qn /norestart
   )
   start /wait "Software uninstallation" /D"%INSTALLER_PATH%" msiexec /x "%INSTALL_CMD%" %OPTIONS% %CUSTOM_OPTIONS%
   set EXIT_CODE=%ERRORLEVEL%
   goto end
   
   
   :install4j
   echo Installing "%INSTALL_CMD%"
   start /wait "Software installation" /D"%INSTALLER_PATH%" "%INSTALL_CMD%" -q %OPTIONS% %CUSTOM_OPTIONS%
   set EXIT_CODE=%ERRORLEVEL%
   goto end
   
   
   :innoinstaller
   echo Installing "%INSTALL_CMD%"
   :: if "%OPTIONS%" == "" (
   :: set OPTIONS=/verysilent /norestart /sp-
   :: )
   start /wait "Software installation" /D"%INSTALLER_PATH%" "%INSTALL_CMD%" /verysilent /norestart /sp- %OPTIONS% %CUSTOM_OPTIONS%
   set EXIT_CODE=%ERRORLEVEL%
   goto end
   
   
   :installshieldinstaller
   echo Installing "%INSTALL_CMD%"
   start /wait "Software installation" /D"%INSTALLER_PATH%" "%INSTALL_CMD%" /s %OPTIONS% %CUSTOM_OPTIONS%
   set EXIT_CODE=%ERRORLEVEL%
   goto end
   
   
   :nsisinstaller
   echo Installing "%INSTALL_CMD%"
   start /wait "Software installation" /D"%INSTALLER_PATH%" "%INSTALL_CMD%" /S %OPTIONS% %CUSTOM_OPTIONS%
   set EXIT_CODE=%ERRORLEVEL%
   goto end
   
   :custominstaller
   if "%OPTIONS%" == "" goto usage
   echo Installing "%INSTALL_CMD%"
   start /wait "Software installation" /D"%INSTALLER_PATH%" "%INSTALL_CMD%" %OPTIONS% %CUSTOM_OPTIONS%
   set EXIT_CODE=%ERRORLEVEL%
   goto end
   
   :usage
   echo Usage:
   echo "%~nx0 <type> <32-bit-installer> <64-bit-installer> [installer-location [custom-options]]"
   echo where type is one of
   echo     msiinstall        Install the given MSI package
   echo     msiuninstall      Uninstall the given MSI package
   echo     msiupdate         Update the given MSP package
   echo     innosetup         Inno setup
   echo     installshield     Install shield
   echo     nsis              Nullsoft install system (NSIS)
   echo     custom            Custom installer - options required in this case
   echo 32-bit-installer      Full file name (including extension) of 32-bit installer
   echo 64-bit-installer      Full file name (including extension) of 64-bit installer
   echo installer-location    Path where the installers are stored
   echo custom-options        Replace the default installer options with the ones given
   exit /B 99
   
   :end
   exit /B %EXIT_CODE%