@call init-path.cmd
echo Creation of base folder...

echo *****************************
MD %datadrv%%datapath%
MD %appsdrv%%appspath%
MD %appsdrv%%appspath%\Installation
MD %appsdrv%%appspath%\Distribution
MD %userdrv%%userpath%
MD %grpdrv%%grppath%
MD %pubdrv%%pubpath%
MD %optdrv%%optpath%
MD %rmprofdrv%%rmprofpath%
MD %tsprofdrv%%tsprofpath%
MD %commundrv%%communpath%
echo *****************************

