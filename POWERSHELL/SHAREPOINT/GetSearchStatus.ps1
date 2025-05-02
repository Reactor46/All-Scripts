Get-SPEnterpriseSearchServiceInstance | select server, status
	 
	Get-SPServiceApplication | ?{$_.name -like "*Search*"}
$spapp = Get-SPServiceApplication -identity [ID]

Get-SPTimerJob | ?{$_.schedule.description -eq "One-time"} |select displayname,server,locktype,lastruntime | fl