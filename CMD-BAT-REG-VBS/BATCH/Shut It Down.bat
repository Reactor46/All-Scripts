ECHO OFF 
TITLE SHUT IT DOWN v0.2.1
CLS
:MENU0
ECHO.
ECHO    _______________________________
ECHO   /                               \
ECHO  [ 1 - Get Help                    ]
ECHO  [ 2 - Go Graphical                ]
ECHO  [ 3 - Remote Control              ]
ECHO  [ 4 - ABORT                       ]
ECHO  [ 5 - More Menus                  ]
ECHO  [ 6 - Changelog                   ]
ECHO  [ 7 - QUIT                        ]  
ECHO   \_______________________________/
ECHO.
SET /P J=TYPE 1, 2, 3, 4, 5, 6, OR 7, THEN PRESS ENTER:
IF %J%==1 GOTO HELP
IF %J%==2 GOTO GUI
IF %J%==3 GOTO RC
IF %J%==4 GOTO ABORT
IF %J%==5 GOTO MENU1
IF %J%==6 GOTO CL
IF %J%==7 GOTO QUIT
IF %J%==11 EXIT
:HELP 
shutdown /? 
PAUSE
GOTO MENU0
:GUI
shutdown /i
PAUSE
GOTO MENU0
:RC
ECHO.
ECHO WHAT WOULD YOU LIKE TO DO?
ECHO.
ECHO :------------------------:
ECHO : 1 - Shutdown           :
ECHO : 2 - Restart            :
ECHO : 3 - Restart with Apps  :
ECHO :------------------------:
ECHO.
REM Replace /s with /r to restart or /g to restart and open registered applications!
SET /P J=TYPE 1, 2, OR 3:
IF %J%==1 GOTO RC1
IF %J%==2 GOTO RC2
IF %J%==3 GOTO RC3
IF %J%==11 EXIT
:RC1
ECHO SHUTDOWN
SET /p M=TARGET:
SET /p T=TIME-OUT:
Shutdown /s /m \\%M% /t %T% 
PAUSE
GOTO RC
:RC2
ECHO RESTART
SET /p M=TARGET:
SET /p T=TIME-OUT:
Shutdown /r /m \\%M% /t %T% 
PAUSE
GOTO RC
:RC3
ECHO RESTART WITH REGISTERED APPS
SET /p M=TARGET:
SET /p T=TIME-OUT:
Shutdown /g /m \\%M% /t %T% 
PAUSE
GOTO RC
:ABORT
shutdown /a
PAUSE
GOTO MENU0
:QUIT
EXIT
:MENU1
ECHO.
ECHO    _______________________________
ECHO   /                               \
ECHO  [ 1 - Instant Shutdown            ]
ECHO  [ 2 - Local Hibernate             ]
ECHO  [ 3 - Close Apps                  ]
ECHO  [ 4 - Close Apps Remote           ]
ECHO  [ 5 - RETURN                      ]
ECHO   \_______________________________/
ECHO.
SET /P J=TYPE 1, 2, 3, 4, OR 5, THEN PRESS ENTER:
IF %J%==1 GOTO INSTANT
IF %J%==2 GOTO LH
IF %J%==3 GOTO CA
IF %J%==4 GOTO CAR
IF %J%==5 GOTO MENU0
IF %J%==11 EXIT
:INSTANT
ECHO Do you REALLY want to do this?
PAUSE
ECHO You asked for it!
shutdown /p
GOTO :EOF
:LH
ECHO Preparing to sleep....
PAUSE
shutdown /h
ECHO Sleeping...
GOTO :EOF
:CA
ECHO Terminate ALL RUNNING apps?
PAUSE
shutdown /f
ECHO TERMINATING.....
GOTO :EOF
:CAR
ECHO.
ECHO :---------------:
ECHO : 1 - Document? :                                  
ECHO : 2 - Comment?  :
ECHO : 3 - Both?     :
ECHO : 4 - RETURN    :
ECHO :---------------:
ECHO.
SET /P J=TYPE 1, 2, 3, OR 4:
IF %J%==1 GOTO CA1
IF %J%==2 GOTO CA2
IF %J%==3 GOTO CA3
IF %J%==4 GOTO MENU1
IF %J%==5 GOTO MENU0
IF %J%==11 EXIT
:CA1
ECHO DOCUMENTING ENABLED
SET /p M=TARGET:
ECHO CONTINUE?
PAUSE
Shutdown  /f /m \\%M% 
ECHO DONE
PAUSE
GOTO CAR
:CA2
ECHO COMMENTS DISABLED
SET /p M=TARGET:
SET /p C=Comment:
ECHO CONTINUE?
PAUSE
Shutdown /f /m \\%M% /c "%C%"
GOTO CAR
:CA3
ECHO DOCUMENTING ENABLED
ECHO COMMENTS ENABLED
SET /p M=TARGET:
SET /p C=Comment:
ECHO CONTINUE?
PAUSE
Shutdown /e /f /m \\%M% /c "%C%" 
GOTO CAR
:CL
ECHO HERE YOU GO!
TYPE Changelog.txt
ECHO WHAT A READ!
PAUSE
GOTO Menu0


