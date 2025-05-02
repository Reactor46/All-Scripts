@ECHO ON
takeown /f KSreenivasan /A /R /D Y 
icacls KSreenivasan /remove ContosoCORP\battistaj /T /C
icacls KSreenivasan /remove ContosoCORP\blatnikj /T /C
icacls KSreenivasan /remove ContosoCORP\leavitts /T /C
icacls KSreenivasan /remove ContosoCORP\bourlandj /T /C
icacls KSreenivasan /grant administrators:F /T /C
icacls KSreenivasan /setowner ContosoCORP\KSreenivasan /T /C
icacls KSreenivasan /inheritance:E /T /C
icacls KSreenivasan /grant ContosoCORP\KSreenivasan:(OI)(CI)F /T