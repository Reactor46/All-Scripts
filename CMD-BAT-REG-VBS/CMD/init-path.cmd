echo Variable Init
echo *****************************

set site=site
echo Location code must be between 2 and 4 letters/number without space or special car.
echo Location code is : %site%

set dnsdom=entreprise.local
echo The DNS domain for this site is : %dnsdom%

set ldapdom=dc=entreprise,dc=local
echo The DNS domain format in LDAP is : %ldapdom%

set srv=%computername%
echo The server name is : %srv%

set ntpsrv=ntp.nrc.ca
echo Internet time server is : %ntpsrv%

set datadrv=d:
set datapath=\Data
echo Root path for data is : %datadrv%%datapath%

set appsdrv=%datadrv%
set appspath=%datapath%\Apps-%site%
echo Root path for Apps is : %appsdrv%%appspath%
 
set userdrv=%datadrv%
set userpath=%datapath%\Home-%site%
echo Root path for users data : %userdrv%%userpath%

set grpdrv=%datadrv%
set grppath=%datapath%\Grp-%site%
echo Root path for group data : %grpdrv%%grppath%

set pubdrv=%datadrv%
set pubpath=%datapath%\Pub-%site%
echo Root path for the public folder : %pubdrv%%pubpath%

set rmprofdrv=%datadrv%
set rmprofpath=%datapath%\RPro-%site%
echo Root path for roaming profile : %rmprofdrv%%rmprofpath%

set tsprofdrv=%datadrv%
set tsprofpath=%datapath%\TPro-%site%
echo Root path for TS profile : %tsprofdrv%%tsprofpath%

set optdrv=%datadrv%
set optpath=\Options\%computername%
echo Root path for optinals build folder : %optdrv%%optpath%

set commundrv=%datadrv%
set communpath=%datapath%\Commun
echo Root path for the commun folder (DFS) : %commundrv%%communpath%

echo *****************************