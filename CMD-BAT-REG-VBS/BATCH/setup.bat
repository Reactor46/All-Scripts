#Setup the task scheduler

schtasks /create /tn "Harvest Job" /tr DailyMailboxSearch.bat /sc weekly /mo 1 /d MON

