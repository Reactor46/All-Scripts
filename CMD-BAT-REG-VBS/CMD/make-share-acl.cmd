@call init-path.cmd
echo Init base share permission

echo *****************************
SetACL \\%srv%\Apps-%site%$ /share /grant S-1-1-0 /full /sid
SetACL \\%srv%\Home-%site%$ /share /grant S-1-1-0 /full /sid
SetACL \\%srv%\Grp-%site%$ /share /grant S-1-1-0 /full /sid	
SetACL \\%srv%\Pub-%site%$ /share /grant S-1-1-0 /full /sid
SetACL \\%srv%\Opt-%site%$ /share /grant S-1-1-0 /full /sid
SetACL \\%srv%\GPO-%site%$ /share /grant S-1-1-0 /full /sid
SetACL \\%srv%\RPro-%site%$ /share /grant S-1-1-0 /full /sid
SetACL \\%srv%\TPro-%site%$ /share /grant S-1-1-0 /full /sid
SetACL \\%srv%\Commun$ /share /grant S-1-1-0 /full /sid
echo *****************************

REM Reminder: setacl must be in the toolkits folder to work.
