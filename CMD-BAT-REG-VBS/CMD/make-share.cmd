@call init-path.cmd
echo Preparing base share permission..

echo *****************************
net share Apps-%site%$=%appsdrv%%appspath%
net share Home-%site%$=%userdrv%%userpath%
net share Grp-%site%$=%grpdrv%%grppath%
net share Pub-%site%$=%pubdrv%%pubpath%
net share Opt-%site%$=%optdrv%%optpath%
net share RPro-%site%$=%rmprofdrv%%rmprofpath%
net share TPro-%site%$=%tsprofdrv%%tsprofpath%
net share Commun$=%commundrv%%communpath%
echo *****************************
pause

