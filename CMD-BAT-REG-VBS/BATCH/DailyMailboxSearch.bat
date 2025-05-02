# Add this script to Windows task scheduler. Modify D:\exchsrvr\bin\exshell.psc1 as 
# per the settings on the exchage server box where this script will be run.

powershell -PSConsoleFile "d:\Program Files\Microsoft\Exchange Server\V14\Bin\exshell.psc1" -command ".\SearchMailboxes.ps1"