@echo off
cls
SetLocal

:: Set your domain name here
set MyDomain=DC=MyDomain,DC=Local
set EMailDomain=mydomain.local

:: Number of Groups, Accounts, & Contacts to create per OU (0 Based)
set NumAccounts=9

:: Set Share Drive
set ShareDrive=F:

:: Set Sub OU's to be created under each catagory. If name has spaced do not put "'s around name.
:: Each Sub OU must start with a differnet character/letter.
::   (this is needed by this script not AD to ensure each account, group, or contact has a unique name)
set SubOU1=North
set SubOU2=South
set SubOU3=East
set SubOU4=West

:: -----------------------------------------------------------------------------------------------------
:: Call :MakeTestOU "MainTestOU" "SubOU1" "SubOU2" "SubOU3" "SubOU4" "SubOU5" "SubOU6" "SubOU7" "SubOU8"
:: "MainTestOU" - Each OU must start with a differnet character/letter.
::   (this is needed by this script not AD to ensure each account, group, or contact has a unique name)
::   SubOUs all must all start with a different letter if they are within the same MainTestOU
:: -----------------------------------------------------------------------------------------------------
Call :MakeTestOU "1 Time Zones" "Eastern" "Central" "Mountain" "Pacific"
Call :MakeTestOU "2 Time Zones" "Eastern" "Central" "Mountain" "Pacific"
Call :MakeTestOU "3 Time Zones" "Eastern" "Central" "Mountain" "Pacific"
Call :MakeTestOU "4 Time Zones" "Eastern" "Central" "Mountain" "Pacific"

goto :eof

:MakeTestOU
SetLocal
set MainTestOU=%~1
if not "MainTestOU"=="" (
  set MTOUFC=%MainTestOU:~0,1%
  dsrm "OU=%MainTestOU%,%MyDomain%" -noprompt -subtree > Nul 2>&1
  call :CreateShares
  call :MakeOUs
  if not "%~2"=="" (
    call :MakeOUs "%~2"
  )
  if not "%~3"=="" (
    call :MakeOUs "%~3"
  )
  if not "%~4"=="" (
    call :MakeOUs "%~4"
  )
  if not "%~5"=="" (
    call :MakeOUs "%~5"
  )
  if not "%~6"=="" (
    call :MakeOUs "%~6"
  )
  if not "%~7"=="" (
    call :MakeOUs "%~7"
  )
  if not "%~8"=="" (
    call :MakeOUs "%~8"
  )
  if not "%~9"=="" (
    call :MakeOUs "%~9"
  )
)
EndLocal
goto :eof

:: --------------------------------------------------
:: Create File Shares
:: --------------------------------------------------
:CreateShares
if exist %ShareDrive%\Shares\HomeDirs (
  net share HomeDir /delete > Nul 2>&1
  rd /s /q %ShareDrive%\Shares\HomeDirs > Nul 2>&1
)
if exist %ShareDrive%\Shares\Profiles (
  net share Profile /delete > Nul 2>&1
  rd /s /q %ShareDrive%\Shares\Profiles > Nul 2>&1
)
md %ShareDrive%\Shares\HomeDirs > Nul 2>&1
net share HomeDir=%ShareDrive%\Shares\HomeDirs /grant:"%UserDomain%\Domain Users",Full > Nul 2>&1
md %ShareDrive%\Shares\Profiles > Nul 2>&1
net share Profile=%ShareDrive%\Shares\Profiles /grant:"%UserDomain%\Domain Users",Full > Nul 2>&1
goto :eof


:: --------------------------------------------------
:: Make Org Units
:: --------------------------------------------------
:MakeOUs
SetLocal
:: Set variables to correct values if function was call with options or not
set OUName=%~1
if not "%OUName%"=="" (
  set FirstChar=%OUName:~0,1%
  set OUName=OU=%OUName%,
) else (
  set FirstChar=
  set OUName=
)
dsadd OU "%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %MainTestOU%"
echo.

:: --------------------------------------------------
:: Create Groups
:: --------------------------------------------------
dsadd OU "OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %MainTestOU%"
echo.
dsadd OU "OU=%SubOU1%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU1%"
dsadd OU "OU=%SubOU2%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU2%"
dsadd OU "OU=%SubOU3%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU3%"
dsadd OU "OU=%SubOU4%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU4%"
echo.
Call :MakeGroups "%~1"
echo.

:: --------------------------------------------------
:: Create Distribution Groups
:: --------------------------------------------------
dsadd OU "OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %MainTestOU%"
echo.
dsadd OU "OU=%SubOU1%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU1%"
dsadd OU "OU=%SubOU2%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU2%"
dsadd OU "OU=%SubOU3%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU3%"
dsadd OU "OU=%SubOU4%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU4%"
echo.
Call :MakeDGroups "%~1"
echo.

:: --------------------------------------------------
:: Create Contacts
:: --------------------------------------------------
dsadd OU "OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %MainTestOU%"
echo.
dsadd OU "OU=%SubOU1%,OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU1%"
dsadd OU "OU=%SubOU2%,OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU2%"
dsadd OU "OU=%SubOU3%,OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU3%"
dsadd OU "OU=%SubOU4%,OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU4%"
echo.
Call :MakeContacts "%~1"
echo.

:: --------------------------------------------------
:: Create Computer Accounts
:: --------------------------------------------------
dsadd OU "OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %MainTestOU%"
echo.
dsadd OU "OU=%SubOU1%,OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU1%"
dsadd OU "OU=%SubOU2%,OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU2%"
dsadd OU "OU=%SubOU3%,OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU3%"
dsadd OU "OU=%SubOU4%,OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU4%"
echo.
Call :MakeComps "%~1"
echo.

:: --------------------------------------------------
:: Create User Accounts
:: --------------------------------------------------
dsadd OU "OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %MainTestOU%"
echo.
dsadd OU "OU=%SubOU1%,OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU1%"
dsadd OU "OU=%SubOU2%,OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU2%"
dsadd OU "OU=%SubOU3%,OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU3%"
dsadd OU "OU=%SubOU4%,OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -desc "Discription %SubOU4%"
echo.
Call :MakeUsers "%~1"
echo.
echo.
EndLocal
goto :eof

:: --------------------------------------------------
:: Make Groups
:: --------------------------------------------------
:MakeGroups
SetLocal
:: Set variables to correct values if function was call with options or not
set OUName=%~1
if not "%OUName%"=="" (
  set FirstChar=%OUName:~0,1%
  set OUName=OU=%OUName%,
) else (
  set FirstChar=
  set OUName=
)
for /l %%C in (0, 1, %NumAccounts%) do (
  if "%OUName%"=="" (
    dsadd Group "CN=%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU1%" -secgrp Yes
    dsadd Group "CN=%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU2%" -secgrp Yes
    dsadd Group "CN=%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU3%" -secgrp Yes
    dsadd Group "CN=%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU4%" -secgrp Yes
  ) else (
    dsadd Group "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU1%" -secgrp Yes
    dsadd Group "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU2%" -secgrp Yes
    dsadd Group "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU3%" -secgrp Yes
    dsadd Group "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %%FirstChar%SubOU4%" -secgrp Yes
  )
)
echo.
EndLocal
goto :eof

:: --------------------------------------------------
:: Make Distribution Groups
:: --------------------------------------------------
:MakeDGroups
SetLocal
:: Set variables to correct values if function was call with options or not
set OUName=%~1
if not "%OUName%"=="" (
  set FirstChar=%OUName:~0,1%
  set OUName=OU=%OUName%,
) else (
  set FirstChar=
  set OUName=
)
for /l %%C in (0, 1, %NumAccounts%) do (
  if "%OUName%"=="" (
    dsadd Group "CN=%SubOU1:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU1%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU1%" -secgrp No -memberof "CN=%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%"
    dsadd Group "CN=%SubOU2:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU2%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU2%" -secgrp No -memberof "CN=%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%"
    dsadd Group "CN=%SubOU3:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU3%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU3%" -secgrp No -memberof "CN=%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%"
    dsadd Group "CN=%SubOU4:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU4%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU4%" -secgrp No -memberof "CN=%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%"
  ) else (
    dsadd Group "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU1%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU1:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU1%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU1%" -secgrp No
    call :AddOther %%C "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU1%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%"
    dsadd Group "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU2%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU2:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU2%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU2%" -secgrp No
    call :AddOther %%C "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU2%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%"
    dsadd Group "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU3%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU3:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU3%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU3%" -secgrp No
    call :AddOther %%C "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU3%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%"
    dsadd Group "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU4%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU4:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU4%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU4%" -secgrp No
    call :AddOther %%C "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU4%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%"
  )
)
echo.
EndLocal
goto :eof

:: --------------------------------------------------
:: Add DGroup to Group
:: --------------------------------------------------
:AddOther
goto :eof
SetLocal
set CurNum=%~1
set CurGroup=%~2
for /l %%O in (0, 1, %NumAccounts%) do (
  if not [%%0]==[%CurNum%] (
    dsmod group "CN=%SubOU1:~0,1%%MTOUFC%Group%%O,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -AddMbr "%CurGroup%"
    dsmod group "CN=%SubOU2:~0,1%%MTOUFC%Group%%O,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -AddMbr "%CurGroup%"
    dsmod group "CN=%SubOU3:~0,1%%MTOUFC%Group%%O,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -AddMbr "%CurGroup%"
    dsmod group "CN=%SubOU4:~0,1%%MTOUFC%Group%%O,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -AddMbr "%CurGroup%"
  )
)
echo.
EndLocal
goto :eof

:: --------------------------------------------------
:: Make Contacts
:: --------------------------------------------------
:MakeContacts
SetLocal
:: Set variables to correct values if function was call with options or not
set OUName=%~1
if not "%OUName%"=="" (
  set FirstChar=%OUName:~0,1%
  set OUName=OU=%OUName%,
) else (
  set FirstChar=
  set OUName=
)
for /l %%C in (0, 1, %NumAccounts%) do (
  if "%OUName%"=="" (
    dsadd Contact "CN=%SubOU1:~0,1%%MTOUFC%Contact%%C,OU=%SubOU1%,OU=_Contacts,OU=%MainTestOU%,%MyDomain%" -fn %SubOU1:~0,1%%MTOUFC%First%%C -mi M -ln %SubOU1:~0,1%%MTOUFC%Last%%C -display "%SubOU1:~0,1%%MTOUFC%Last%%C, %SubOU1:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU1%" -email %SubOU1:~0,1%%MTOUFC%Contact%%C@%EMailDomain% -office "Office %SubOU1%" -title "Job Title" -dept "Department %SubOU1%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212"
    dsadd Contact "CN=%SubOU2:~0,1%%MTOUFC%Contact%%C,OU=%SubOU2%,OU=_Contacts,OU=%MainTestOU%,%MyDomain%" -fn %SubOU2:~0,1%%MTOUFC%First%%C -mi M -ln %SubOU2:~0,1%%MTOUFC%Last%%C -display "%SubOU2:~0,1%%MTOUFC%Last%%C, %SubOU2:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU2%" -email %SubOU2:~0,1%%MTOUFC%Contact%%C@%EMailDomain% -office "Office %SubOU2%" -title "Job Title" -dept "Department %SubOU2%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212"
    dsadd Contact "CN=%SubOU3:~0,1%%MTOUFC%Contact%%C,OU=%SubOU3%,OU=_Contacts,OU=%MainTestOU%,%MyDomain%" -fn %SubOU3:~0,1%%MTOUFC%First%%C -mi M -ln %SubOU3:~0,1%%MTOUFC%Last%%C -display "%SubOU3:~0,1%%MTOUFC%Last%%C, %SubOU3:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU3%" -email %SubOU3:~0,1%%MTOUFC%Contact%%C@%EMailDomain% -office "Office %SubOU3%" -title "Job Title" -dept "Department %SubOU3%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212"
    dsadd Contact "CN=%SubOU4:~0,1%%MTOUFC%Contact%%C,OU=%SubOU4%,OU=_Contacts,OU=%MainTestOU%,%MyDomain%" -fn %SubOU4:~0,1%%MTOUFC%First%%C -mi M -ln %SubOU4:~0,1%%MTOUFC%Last%%C -display "%SubOU4:~0,1%%MTOUFC%Last%%C, %SubOU4:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU4%" -email %SubOU4:~0,1%%MTOUFC%Contact%%C@%EMailDomain% -office "Office %SubOU4%" -title "Job Title" -dept "Department %SubOU4%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212"
  ) else (
    dsadd Contact "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%Contact%%C,OU=%SubOU1%,OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -fn %FirstChar%%SubOU1:~0,1%%MTOUFC%First%%C -mi M -ln %FirstChar%%SubOU1:~0,1%%MTOUFC%Last%%C -display "%FirstChar%%SubOU1:~0,1%%MTOUFC%Last%%C, %FirstChar%%SubOU1:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU1%" -email %FirstChar%%SubOU1:~0,1%%MTOUFC%Contact%%C@%EMailDomain% -office "Office %SubOU1%" -title "Job Title" -dept "Department %SubOU1%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212"
    dsadd Contact "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%Contact%%C,OU=%SubOU2%,OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -fn %FirstChar%%SubOU2:~0,1%%MTOUFC%First%%C -mi M -ln %FirstChar%%SubOU2:~0,1%%MTOUFC%Last%%C -display "%FirstChar%%SubOU2:~0,1%%MTOUFC%Last%%C, %FirstChar%%SubOU2:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU2%" -email %FirstChar%%SubOU2:~0,1%%MTOUFC%Contact%%C@%EMailDomain% -office "Office %SubOU2%" -title "Job Title" -dept "Department %SubOU2%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212"
    dsadd Contact "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%Contact%%C,OU=%SubOU3%,OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -fn %FirstChar%%SubOU3:~0,1%%MTOUFC%First%%C -mi M -ln %FirstChar%%SubOU3:~0,1%%MTOUFC%Last%%C -display "%FirstChar%%SubOU3:~0,1%%MTOUFC%Last%%C, %FirstChar%%SubOU3:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU3%" -email %FirstChar%%SubOU3:~0,1%%MTOUFC%Contact%%C@%EMailDomain% -office "Office %SubOU3%" -title "Job Title" -dept "Department %SubOU3%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212"
    dsadd Contact "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%Contact%%C,OU=%SubOU4%,OU=_Contacts,%OUName%OU=%MainTestOU%,%MyDomain%" -fn %FirstChar%%SubOU4:~0,1%%MTOUFC%First%%C -mi M -ln %FirstChar%%SubOU4:~0,1%%MTOUFC%Last%%C -display "%FirstChar%%SubOU4:~0,1%%MTOUFC%Last%%C, %FirstChar%%SubOU4:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU4%" -email %FirstChar%%SubOU4:~0,1%%MTOUFC%Contact%%C@%EMailDomain% -office "Office %SubOU4%" -title "Job Title" -dept "Department %SubOU4%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212"
  )
)
echo.
EndLocal
goto :eof

:: --------------------------------------------------
:: Make Computer Accounts
:: --------------------------------------------------
:MakeComps
SetLocal
:: Set variables to correct values if function was call with options or not
set OUName=%~1
if not "%OUName%"=="" (
  set FirstChar=%OUName:~0,1%
  set OUName=OU=%OUName%,
) else (
  set FirstChar=
  set OUName=
)
for /l %%C in (0, 1, %NumAccounts%) do (
  if "%OUName%"=="" (
    dsadd Computer "CN=%SubOU1:~0,1%%MTOUFC%Computer%%C,OU=%SubOU1%,OU=_Computers,OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU1%" -loc "Location %SubOU1%"
    dsadd Computer "CN=%SubOU2:~0,1%%MTOUFC%Computer%%C,OU=%SubOU2%,OU=_Computers,OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU2%" -loc "Location %SubOU2%"
    dsadd Computer "CN=%SubOU3:~0,1%%MTOUFC%Computer%%C,OU=%SubOU3%,OU=_Computers,OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU3%" -loc "Location %SubOU3%"
    dsadd Computer "CN=%SubOU4:~0,1%%MTOUFC%Computer%%C,OU=%SubOU4%,OU=_Computers,OU=%MainTestOU%,%MyDomain%" -memberof "CN=%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %SubOU4%" -loc "Location %SubOU4%"
  ) else (
    dsadd Computer "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%Computer%%C,OU=%SubOU1%,OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU1%" -loc "Location %FirstChar%%SubOU1%"
    dsadd Computer "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%Computer%%C,OU=%SubOU2%,OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU2%" -loc "Location %FirstChar%%SubOU2%"
    dsadd Computer "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%Computer%%C,OU=%SubOU3%,OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU3%" -loc "Location %FirstChar%%SubOU3%"
    dsadd Computer "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%Computer%%C,OU=%SubOU4%,OU=_Computers,%OUName%OU=%MainTestOU%,%MyDomain%" -memberof "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" -desc "Description %FirstChar%%SubOU4%" -loc "Location %FirstChar%%SubOU4%"
  )
)
echo.
EndLocal
goto :eof

:: --------------------------------------------------
:: Make User Accounts
:: --------------------------------------------------
:MakeUsers
SetLocal
:: Set variables to correct values if function was call with options or not
set OUName=%~1
if not "%OUName%"=="" (
  set FirstChar=%OUName:~0,1%
  set OUName=OU=%OUName%,
) else (
  set FirstChar=
  set OUName=
)
for /l %%C in (0, 1, %NumAccounts%) do (
  if "%OUName%"=="" (
    dsadd User "CN=%SubOU1:~0,1%%MTOUFC%User%%C,OU=%SubOU1%,OU=_Users,OU=%MainTestOU%,%MyDomain%" -samid %SubOU1:~0,1%%MTOUFC%User%%C -upn %SubOU1:~0,1%%MTOUFC%User%%C@%EMailDomain% -empid %SubOU1:~0,1%%MTOUFC%User%%C -pwd "Us3rP@ssw0rd" -mustchpwd no -disabled no -memberof "CN=%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%SubOU1:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU1%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -fn %SubOU1:~0,1%%MTOUFC%First%%C -mi M -ln %SubOU1:~0,1%%MTOUFC%Last%%C -display "%SubOU1:~0,1%%MTOUFC%Last%%C, %SubOU1:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU1%" -email %SubOU1:~0,1%%MTOUFC%User%%C@%EMailDomain% -office "Office %SubOU1%" -title "Job Title" -dept "Department %SubOU1%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212" -webpg http://%EMailDomain%/%SubOU1:~0,1%%MTOUFC%User%%C -hmdir \\%ComputerName%\HomeDir\%SubOU1:~0,1%%MTOUFC%User%%C -hmdrv H: -profile \\%ComputerName%\Profile\%SubOU1:~0,1%%MTOUFC%User%%C -loscr LogonScript.cmd
    dsadd User "CN=%SubOU2:~0,1%%MTOUFC%User%%C,OU=%SubOU2%,OU=_Users,OU=%MainTestOU%,%MyDomain%" -samid %SubOU2:~0,1%%MTOUFC%User%%C -upn %SubOU2:~0,1%%MTOUFC%User%%C@%EMailDomain% -empid %SubOU2:~0,1%%MTOUFC%User%%C -pwd "Us3rP@ssw0rd" -mustchpwd no -disabled no -memberof "CN=%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%SubOU2:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU2%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -fn %SubOU2:~0,1%%MTOUFC%First%%C -mi M -ln %SubOU2:~0,1%%MTOUFC%Last%%C -display "%SubOU2:~0,1%%MTOUFC%Last%%C, %SubOU2:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU2%" -email %SubOU2:~0,1%%MTOUFC%User%%C@%EMailDomain% -office "Office %SubOU2%" -title "Job Title" -dept "Department %SubOU2%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212" -webpg http://%EMailDomain%/%SubOU2:~0,1%%MTOUFC%User%%C -hmdir \\%ComputerName%\HomeDir\%SubOU2:~0,1%%MTOUFC%User%%C -hmdrv H: -profile \\%ComputerName%\Profile\%SubOU2:~0,1%%MTOUFC%User%%C -loscr LogonScript.cmd
    dsadd User "CN=%SubOU3:~0,1%%MTOUFC%User%%C,OU=%SubOU3%,OU=_Users,OU=%MainTestOU%,%MyDomain%" -samid %SubOU3:~0,1%%MTOUFC%User%%C -upn %SubOU3:~0,1%%MTOUFC%User%%C@%EMailDomain% -empid %SubOU3:~0,1%%MTOUFC%User%%C -pwd "Us3rP@ssw0rd" -mustchpwd no -disabled no -memberof "CN=%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%SubOU3:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU3%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -fn %SubOU3:~0,1%%MTOUFC%First%%C -mi M -ln %SubOU3:~0,1%%MTOUFC%Last%%C -display "%SubOU3:~0,1%%MTOUFC%Last%%C, %SubOU3:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU3%" -email %SubOU3:~0,1%%MTOUFC%User%%C@%EMailDomain% -office "Office %SubOU3%" -title "Job Title" -dept "Department %SubOU3%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212" -webpg http://%EMailDomain%/%SubOU3:~0,1%%MTOUFC%User%%C -hmdir \\%ComputerName%\HomeDir\%SubOU3:~0,1%%MTOUFC%User%%C -hmdrv H: -profile \\%ComputerName%\Profile\%SubOU3:~0,1%%MTOUFC%User%%C -loscr LogonScript.cmd
    dsadd User "CN=%SubOU4:~0,1%%MTOUFC%User%%C,OU=%SubOU4%,OU=_Users,OU=%MainTestOU%,%MyDomain%" -samid %SubOU4:~0,1%%MTOUFC%User%%C -upn %SubOU4:~0,1%%MTOUFC%User%%C@%EMailDomain% -empid %SubOU4:~0,1%%MTOUFC%User%%C -pwd "Us3rP@ssw0rd" -mustchpwd no -disabled no -memberof "CN=%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%SubOU4:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU4%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -fn %SubOU4:~0,1%%MTOUFC%First%%C -mi M -ln %SubOU4:~0,1%%MTOUFC%Last%%C -display "%SubOU4:~0,1%%MTOUFC%Last%%C, %SubOU4:~0,1%%MTOUFC%First%%C M" -desc "Description %SubOU4%" -email %SubOU4:~0,1%%MTOUFC%User%%C@%EMailDomain% -office "Office %SubOU4%" -title "Job Title" -dept "Department %SubOU4%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212" -webpg http://%EMailDomain%/%SubOU4:~0,1%%MTOUFC%User%%C -hmdir \\%ComputerName%\HomeDir\%SubOU4:~0,1%%MTOUFC%User%%C -hmdrv H: -profile \\%ComputerName%\Profile\%SubOU4:~0,1%%MTOUFC%User%%C -loscr LogonScript.cmd
  ) else (
    dsadd User "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%User%%C,OU=%SubOU1%,OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -samid %FirstChar%%SubOU1:~0,1%%MTOUFC%User%%C -upn %FirstChar%%SubOU1:~0,1%%MTOUFC%User%%C@%EMailDomain% -empid %FirstChar%%SubOU1:~0,1%%MTOUFC%User%%C -pwd "Us3rP@ssw0rd" -mustchpwd no -disabled no -memberof "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU1:~0,1%%MTOUFC%Group%%C,OU=%SubOU1%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU4%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU2:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU2%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -fn %FirstChar%%SubOU1:~0,1%%MTOUFC%First%%C -mi M -ln %FirstChar%%SubOU1:~0,1%%MTOUFC%Last%%C -display "%FirstChar%%SubOU1:~0,1%%MTOUFC%Last%%C, %FirstChar%%SubOU1:~0,1%%MTOUFC%First%%C M" -desc "Description %FirstChar%%SubOU1%" -email %FirstChar%%SubOU1:~0,1%%MTOUFC%User%%C@%EMailDomain% -office "Office %FirstChar%%SubOU1%" -title "Job Title" -dept "Department %FirstChar%%SubOU1%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212" -mgr "CN=%SubOU1:~0,1%%MTOUFC%User%%C,OU=%SubOU1%,OU=_Users,OU=%MainTestOU%,%MyDomain%" -webpg http://%EMailDomain%/%FirstChar%%SubOU1:~0,1%%MTOUFC%User%%C -hmdir \\%ComputerName%\HomeDir\%FirstChar%%SubOU1:~0,1%%MTOUFC%User%%C -hmdrv H: -profile \\%ComputerName%\Profile\%FirstChar%%SubOU1:~0,1%%MTOUFC%User%%C -loscr LogonScript.cmd
    dsadd User "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%User%%C,OU=%SubOU2%,OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -samid %FirstChar%%SubOU2:~0,1%%MTOUFC%User%%C -upn %FirstChar%%SubOU2:~0,1%%MTOUFC%User%%C@%EMailDomain% -empid %FirstChar%%SubOU2:~0,1%%MTOUFC%User%%C -pwd "Us3rP@ssw0rd" -mustchpwd no -disabled no -memberof "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU2:~0,1%%MTOUFC%Group%%C,OU=%SubOU2%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU3%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU1:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU1%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -fn %FirstChar%%SubOU2:~0,1%%MTOUFC%First%%C -mi M -ln %FirstChar%%SubOU2:~0,1%%MTOUFC%Last%%C -display "%FirstChar%%SubOU2:~0,1%%MTOUFC%Last%%C, %FirstChar%%SubOU2:~0,1%%MTOUFC%First%%C M" -desc "Description %FirstChar%%SubOU2%" -email %FirstChar%%SubOU2:~0,1%%MTOUFC%User%%C@%EMailDomain% -office "Office %FirstChar%%SubOU2%" -title "Job Title" -dept "Department %FirstChar%%SubOU2%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212" -mgr "CN=%SubOU2:~0,1%%MTOUFC%User%%C,OU=%SubOU2%,OU=_Users,OU=%MainTestOU%,%MyDomain%" -webpg http://%EMailDomain%/%FirstChar%%SubOU2:~0,1%%MTOUFC%User%%C -hmdir \\%ComputerName%\HomeDir\%FirstChar%%SubOU2:~0,1%%MTOUFC%User%%C -hmdrv H: -profile \\%ComputerName%\Profile\%FirstChar%%SubOU2:~0,1%%MTOUFC%User%%C -loscr LogonScript.cmd
    dsadd User "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%User%%C,OU=%SubOU3%,OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -samid %FirstChar%%SubOU3:~0,1%%MTOUFC%User%%C -upn %FirstChar%%SubOU3:~0,1%%MTOUFC%User%%C@%EMailDomain% -empid %FirstChar%%SubOU3:~0,1%%MTOUFC%User%%C -pwd "Us3rP@ssw0rd" -mustchpwd no -disabled no -memberof "CN=%FirstChar%%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU3:~0,1%%MTOUFC%Group%%C,OU=%SubOU3%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%FirstChar%%SubOU2:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU2%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU4:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU4%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -fn %FirstChar%%SubOU3:~0,1%%MTOUFC%First%%C -mi M -ln %FirstChar%%SubOU3:~0,1%%MTOUFC%Last%%C -display "%FirstChar%%SubOU3:~0,1%%MTOUFC%Last%%C, %FirstChar%%SubOU3:~0,1%%MTOUFC%First%%C M" -desc "Description %FirstChar%%SubOU3%" -email %FirstChar%%SubOU3:~0,1%%MTOUFC%User%%C@%EMailDomain% -office "Office %FirstChar%%SubOU3%" -title "Job Title" -dept "Department %FirstChar%%SubOU3%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212" -mgr "CN=%SubOU3:~0,1%%MTOUFC%User%%C,OU=%SubOU3%,OU=_Users,OU=%MainTestOU%,%MyDomain%" -webpg http://%EMailDomain%/%FirstChar%%SubOU3:~0,1%%MTOUFC%User%%C -hmdir \\%ComputerName%\HomeDir\%FirstChar%%SubOU3:~0,1%%MTOUFC%User%%C -hmdrv H: -profile \\%ComputerName%\Profile\%FirstChar%%SubOU3:~0,1%%MTOUFC%User%%C -loscr LogonScript.cmd
    dsadd User "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%User%%C,OU=%SubOU4%,OU=_Users,%OUName%OU=%MainTestOU%,%MyDomain%" -samid %FirstChar%%SubOU4:~0,1%%MTOUFC%User%%C -upn %FirstChar%%SubOU4:~0,1%%MTOUFC%User%%C@%EMailDomain% -empid %FirstChar%%SubOU4:~0,1%%MTOUFC%User%%C -pwd "Us3rP@ssw0rd" -mustchpwd no -disabled no -memberof "CN=%FirstChar%%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU4:~0,1%%MTOUFC%Group%%C,OU=%SubOU4%,OU=_Groups,OU=%MainTestOU%,%MyDomain%" "CN=%FirstChar%%SubOU1:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU1%,OU=_DGroups,%OUName%OU=%MainTestOU%,%MyDomain%" "CN=%SubOU3:~0,1%%MTOUFC%DGroup%%C,OU=%SubOU3%,OU=_DGroups,OU=%MainTestOU%,%MyDomain%" -fn %FirstChar%%SubOU4:~0,1%%MTOUFC%First%%C -mi M -ln %FirstChar%%SubOU4:~0,1%%MTOUFC%Last%%C -display "%FirstChar%%SubOU4:~0,1%%MTOUFC%Last%%C, %FirstChar%%SubOU4:~0,1%%MTOUFC%First%%C M" -desc "Description %FirstChar%%SubOU4%" -email %FirstChar%%SubOU4:~0,1%%MTOUFC%User%%C@%EMailDomain% -office "Office %FirstChar%%SubOU4%" -title "Job Title" -dept "Department %FirstChar%%SubOU4%" -company "Company %EMailDomain%" -tel "(555) 555-1212" -hometel "(555) 555-1212" -pager "(555) 555-1212" -mobile "(555) 555-1212" -fax "(555) 555-1212" -mgr "CN=%SubOU4:~0,1%%MTOUFC%User%%C,OU=%SubOU4%,OU=_Users,OU=%MainTestOU%,%MyDomain%" -webpg http://%EMailDomain%/%FirstChar%%SubOU4:~0,1%%MTOUFC%User%%C -hmdir \\%ComputerName%\HomeDir\%FirstChar%%SubOU4:~0,1%%MTOUFC%User%%C -hmdrv H: -profile \\%ComputerName%\Profile\%FirstChar%%SubOU4:~0,1%%MTOUFC%User%%C -loscr LogonScript.cmd
  )
)
echo.
EndLocal
goto :eof
