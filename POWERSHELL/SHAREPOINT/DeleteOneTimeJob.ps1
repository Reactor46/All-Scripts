$tjs = Get-SPTimerJob | ?{$_.schedule.description -eq "One-time"}  
foreach ($tj in $tjs) 
{$tj.delete()}