REM - Save this as c:\scripts\tl.bat (save the bottom REM lines (without the 1st REM) as c:\scripts\stripit.bat
REM - c:\scripts\tl.bat - List all schedule tasks on all remote (NON-domain-controller) servers
REM - NOTE 1: Does NOT list scheduled tasks on Domain Controllers
REM - NOTE 2: Use "stripit.bat" also to exclude special servers (i.e., 'TestServer1$', cluster aliases, etc.)
REM - NOTE 3: Change YOUR-DOMAIN info to match your own domain, as well as any 'OU' you wish to use (i.e. "Servers")
REM - Author: Jeff Mason / aka 'bitdoctor' aka 'tnjman' / 06/08/2013 [willspy4u]\H0+ma1l/c*m

if exist c:\scripts\server_raw.txt del c:\scripts\server_raw.txt
if exist c:\scripts\server_list.txt del c:\scripts\server_list.txt

dsquery computer ou="Servers",dc=YOUR-DOMAIN,dc=COM | dsget computer -samid -c > c:\scripts\server_raw.txt

REM - [we CALL stripit.bat] strips "$" from end of server name in "server_raw.txt" & creates CLEAN "server_list.txt"

for /f %%a in (c:\scripts\server_raw.txt) do call c:\scripts\stripit.bat %%a

for /f %%a in (c:\scripts\server_list.txt) do schtasks /query /s %%a /v /fo csv >> c:\scripts\sched_tasks.csv

goto fin

REM - You can use below lines to make a stripit.bat file (else turn below into subroutine)
REM - Take out the FIRST "REM" in each line below, and save it as c:\scripts\stripit.bat (first 2 lines skip the "dsget" and "samid" lines)

REM REM c:\scripts\stripit.bat
REM if %1%==TestServer1$ goto skipit
REM if %1%==dsget goto skipit
REM if %1%==samid goto skipit
REM set sname=%1
REM set sname=%sname:~0,-1%
REM echo %sname% >> c:\scripts\server_list.txt
REM :skipit

:fin