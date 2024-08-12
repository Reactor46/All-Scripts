:: "Search for duplicate lines"
:: forums.petri.com/showthread.php?t=32793
:: author: Remco Simons [NL] 2009

@echo off
Setlocal ENABLEDELAYEDEXPANSION

::# Search for duplicate lines in:
Set "TXTFile=c:\LazyWinAdmin\sid_Dupe.txt"

echo/%TXTFile%
echo/                  (Empy lines are not being counted!) &echo/
echo/------------------------------------------------------------------------------+

title Find duplicate lines:
Set "skipLines=,"
For /f "usebackq delims=" %%! in ("%TXTFile%") do (
  Set/a lnCnt2=0
  Set/a iCnt3=0
  Set "fndLines="
  Set "doubleLine="
  Set/a lnCnt1=!lnCnt1!+1
  Set "readline=%%!"

  For /f "usebackq delims=" %%* in ("%TXTFile%") do (
    Set/a lnCnt2=!lnCnt2!+1
    If !lnCnt1! LSS !lnCnt2! (
      Set/a "l=10000+!lnCnt2!" & Set "l=!l:~1!"
      (echo/!skipLines! |Find /v ",!l!,")>nul &&(
        Set "compareline=%%*"
        If /i "!readline!" EQU "!compareline!" (
          Set/a iCnt3=!iCnt3!+1
          Set "skipLines=!skipLines:~0,-1!,!l!,"
          Set "fndLines=!fndLines!, !l!"
          Set "doubleLine=!compareline!
        )
      )
    )
  )

  If !iCnt3! GTR 0 (
    Set "fndLines=!fndLines:~2!
    ECHO/Line !lnCnt1!:
    ECHO/"!doubleLine!"
    ECHO/, the same line was found !iCnt3! more times
    ECHO/  at the line(s^): !fndLines!
    echo/------------------------------------------------------------------------------+
  )
)

echo/&echo/Done & pause>nul