@ECHO ON
 takeown /f lachavez /A /R /D Y 
 icacls lachavez /ove FNBMCORP\battistaj /T /C
 icacls lachavez /ove FNBMCORP\blatnikj /T /C
 icacls lachavez /ove FNBMCORP\leavitts /T /C
 icacls lachavez /ove FNBMCORP\bourlandj /T /C
 icacls lachavez /grant administrators:F /T /C
 icacls lachavez /setowner FNBMCORP\lchavez /T /C
 icacls lachavez /inheritance:E /T /C
 icacls lachavez /grant FNBMCORP\lchavez:(OI)(CI)F /T