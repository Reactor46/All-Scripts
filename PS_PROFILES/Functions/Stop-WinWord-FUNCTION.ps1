Function Stop-WinWord
{
	Write-Debug "***Enter Stop-WinWord"
	
	## determine our login session
	$proc = Get-Process -PID $PID
	If( $null -eq $proc )
	{
		throw "Stop-WinWord: Cannot find process $PID"
	}
	
	$SessionID = $proc.SessionId
	If( $null -eq $SessionID )
	{
		Write-Debug "Stop-WinWord: SessionId on $PID is null"
		throw "Can't find a session for pid $PID"
	}

	If( 0 -eq $SessionID )
	{
		Write-Debug "Stop-WinWord: SessionId is 0 -- that is a bug"
		throw "SessionId is zero for pid $PID"
	}
	
	#Find out if winword is running in our session
	try 
	{
		$wordProc = Get-Process 'WinWord' -ErrorAction SilentlyContinue
	}
	catch
	{
		Write-Debug "***Exit Stop-WinWord: no WinWord tasks are running #1"
		Return ## not running
	}

	If( !$wordproc )
	{
		Write-Debug "***Exit Stop-WinWord: no WinWord tasks are running #2"
		Return ## WinWord is not running in ANY session
	}
	
	$wordrunning = $wordProc |? { $_.SessionId -eq $SessionID }
	If( !$wordrunning )
	{
		Write-Debug "***Exit Stop-WinWord: wordRunning eq null"
		Return ## not running in the current session
	}
	If( $wordrunning -is [Array] )
	{
		Write-Debug "***Exit Stop-WinWord: wordRunning is an array, elements=$($wordrunning.Count)"
		throw "Multiple Word processes are running in session $SessionID"
	}

	## it is possible for the below to throw a fault if Winword stops before it is executed.
	Stop-Process -Id $wordrunning.Id -ErrorAction SilentlyContinue
	Write-Debug "***Exit Stop-WinWord: sent Stop-Process to $($wordrunning.Id)"
}
