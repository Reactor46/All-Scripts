Function AbortScript
{
	$Word.Quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject( $Word ) | Out-Null
	If( Get-Variable -Name Word -Scope Global )
	{
		Remove-Variable -Name word -Scope Global
	}
	[GC]::Collect() 
	[GC]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}