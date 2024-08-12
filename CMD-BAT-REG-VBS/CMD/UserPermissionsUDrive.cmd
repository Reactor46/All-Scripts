@ECHO ON
 takeown /f lachavez /A /R /D Y 
 icacls lachavez /ove ContosoCORP\battistaj /T /C
 icacls lachavez /ove ContosoCORP\blatnikj /T /C
 icacls lachavez /ove ContosoCORP\leavitts /T /C
 icacls lachavez /ove ContosoCORP\bourlandj /T /C
 icacls lachavez /grant administrators:F /T /C
 icacls lachavez /setowner ContosoCORP\lchavez /T /C
 icacls lachavez /inheritance:E /T /C
 icacls lachavez /grant ContosoCORP\lchavez:(OI)(CI)F /T