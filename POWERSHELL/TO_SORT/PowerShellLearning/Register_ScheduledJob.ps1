#Creating a trigger and saying when the job is meant to take place
$Trigger = New-JobTrigger -Daily -At 4pm

#Defining the script block and saying where the script that actually runs the job is. In this case, the script test my connection to YouTube
$Scriptblock = { C:\Users\silad\Documents\Powershell_stuff\PowerShellLearning\Test_Youtube.ps1 }

#Registering the job and explaining which trigger and scriptblock I'm employing

Register-ScheduledJob -Name 'Test_Youtube' -ScriptBlock $Scriptblock -Trigger $Trigger