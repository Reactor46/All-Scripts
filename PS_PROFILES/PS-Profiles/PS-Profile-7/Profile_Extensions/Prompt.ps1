Function prompt
	{
	if ( $host.name -eq "ConsoleHost" )
		{
		$date = (get-date -format 'ddd MMM dd')
		$time = (get-date -format ' hh:mm:ss tt')
		$location = ((get-location).path).tostring()
		$location = $location.replace('Microsoft.PowerShell.Core\FileSystem::','')
		$location = $location.replace('misfs1\wreeves','NetHome')
		$location = $location.replace('missvr2\mis','MISShare')
		$location = $location.replace('misfs1\chickerson','ChadNet')
		$location = $location.replace('admin9491\b$\chickerson\Documents\WindowsPowershell\scripts','ChadScripts')
		$location = $location.replace('admin9491\b$\chickerson\Documents\WayneShares','ChadDump')
		$location = $location.replace('admin9491\b$\chickerson','ChadHome')
		$location = $location.replace('Users\wreeves','~')
		$location = $location.replace('Documents\WindowsPowershell\Scripts','Scripts')
		$location = $location.replace('Documents\WindowsPowershell','PSHome')
		$location = $location.replace('Scripts\Security','Security')
		#$location = $location.replace('\','/')
		#$location = $location.replace('C:','')
		Write-Host $date -f Green -nonewline;write-host $time -f Yellow -nonewline;write-host " $location>" -f Cyan -NoNewline
		return " "
		}
	}