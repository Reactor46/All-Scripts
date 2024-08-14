@echo off
rem ------------------------------------------------------------------------------------------------------------------------------------
rem - Script:  setpwr.bat
rem - Date:    08/26/2013
rem - Author:  Jeff Mason / aka bitdoctor / aka tnjman
rem - Purpose: To set power management settings on a remote computer
rem - Assumptions: 1) You have 'psexec' located in "c:\tools" folder
rem -              2) You copy this script to a central server, into a scripts share (For example, into: \\script-server\scripts)
rem -
rem - How to run: c:\tools\psexec \\remote-pc -u domain\privileged-ID -p password \\script-server\scripts\setpwr.bat
rem -  Customize the above command for your environment, where:
rem      "remote-pc" is the remote computer whose power management settings you want to change & 
REM      domain\privileged-ID & password are privileged ID/password in the domain that have workstation admin & network rights to 
REM      run the script and set power management settings.
rem
rem -  Example: c:\tools\psexec \\PC101 -u MYDOM\sysadmin2 -p **admin-password** \\SCRIPTSRV1\scripts\setpwr.bat
rem -  The above would set power management settings on PC101
rem - In our situation, we are changing the settings of the "Balanced" power mode (381b4222-f694-41f0-9685-ff5bb260df2e)
rem ------------------------------------------------------------------------------------------------------------------------------------
echo.
echo "Setting variables..." 
set pf=powercfg
set av=setacvalueindex
set ps="UNKNOWN"
set pm=381b4222-f694-41f0-9685-ff5bb260df2e
rem
if %pm%==381b4222-f694-41f0-9685-ff5bb260df2e (
  set ps="Balanced"
)
if %pm%==8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c (
  set ps="High performance"
)
if %pm%==a1841308-3541-4fab-bc81-f71556f20b4a (
  set ps="Power saver"
)
echo.
echo "Power Scheme is %ps%"
echo.
echo.
echo "Changing power configuration settings for power scheme: %ps% on Computer: %computername%"
echo.
rem -- set disk powerdown for 20 min (alternate form: powercfg -x disk-timeout-ac 20)
rem -      scheme, hard disk,                           powerdown after                      1200 = (seconds = 20 mins)
%pf% /%av% %pm%    0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e 1200
rem -      scheme, sleep,                               sleep after                          0 = NEVER SLEEP
%pf% /%av% %pm%    238c9fa8-0aad-41ed-83f4-97be242c8f20 29f6c1db-86da-48c5-9fdb-f2b67b1f44da 0
rem -      scheme, sleep,                               allow hybrid sleep                   0 = OFF (Don't allow hybrid)
%pf% /%av% %pm%    238c9fa8-0aad-41ed-83f4-97be242c8f20 94ac6d29-73ce-41a6-809f-6363ba21b47e 0
rem -      scheme, sleep,                               hibernate after                      0 = NEVER HIBERNATE
%pf% /%av% %pm%    238c9fa8-0aad-41ed-83f4-97be242c8f20 9d7815a6-7ee4-497e-8888-515a05f02364 0
rem -      scheme, sleep,                               allow wake timers                    1 = Enabled
%pf% /%av% %pm%    238c9fa8-0aad-41ed-83f4-97be242c8f20 bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d 1
rem -      scheme, USB,                                 USB selective suspend                0 = Disabled (Don't allow selective suspend)
%pf% /%av% %pm%    2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0
rem -      scheme, PCI Express,                         Link state pwr management            0 = Off
%pf% /%av% %pm%    501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 0
rem -      scheme, Processor power management,          Minimum processor state              100 = Max
%pf% /%av% %pm%    54533251-82be-4824-96c1-47b60b740d00 893dee8e-2bef-41e0-89c6-b55d0929964c 100
rem -      scheme, Processor power management,          Maximum processor state              100 = Max
%pf% /%av% %pm%    54533251-82be-4824-96c1-47b60b740d00 bc5038f7-23e0-4960-96da-33abaf5935ec 100
rem -      scheme, Display,                             Dim Display after                    300 = (seconds = 5 minutes)
%pf% /%av% %pm%    7516b95f-f776-4464-8c53-06167f40cc99 17aaa29b-8b43-4b94-aafe-35f64daaf1ee 300
rem -      scheme, Display,                             Turn off Display after               600 = (seconds = 10 minutes)
%pf% /%av% %pm%    7516b95f-f776-4464-8c53-06167f40cc99 3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e 600
rem -      scheme, Multimedia settings,                 When sharing media                   1 = Pevent idling to sleep
%pf% /%av% %pm%    9596fb26-9850-41fd-ac3e-f7c3c00afd4b 03680956-93bc-4294-bba6-4e0f09bb717f 1
rem -      scheme, Multimedia settings,                 When playing video                   0 = Optimize video quality
%pf% /%av% %pm%    9596fb26-9850-41fd-ac3e-f7c3c00afd4b 34c7b99f-9a6d-4b3c-8dc7-b6693b78cef4 0
rem -      scheme, Internet Explorer,                   JavaScript Timer Frequency           1 = Maximize performance
%pf% /%av% %pm%    b14a8f96-7b67-4e78-8192-b890b1a62b8a 4c793e7d-a264-42e1-87d3-7a0d2f523ccd 1
rem
echo "Finished changing power configuration settings for power scheme: %ps% on Computer: %computername%"
echo.
echo "Setting power scheme %ps% to active mode for computer: %computername%"
rem - Set balanced plan as active
%pf% -setactive %pm%
echo.
echo "Done setting power scheme %ps% to active mode for computer: %computername%. Now exiting."
exit