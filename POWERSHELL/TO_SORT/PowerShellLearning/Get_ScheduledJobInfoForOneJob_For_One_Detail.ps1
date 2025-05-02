#Filtering the Job info so that it just shows if the job is enabled or not
Get-ScheduledJob -Name Test_Youtube | select Enabled