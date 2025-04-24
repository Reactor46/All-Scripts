
function UpgradeSite($path){
 
	$upgradeSessionInfo = Get-SPSiteUpgradeSessionInfo -Site $path
 
	Write-host "Beginning upgrade process for"$path  -foregroundcolor black -backgroundcolor green
 
	#queue the site for upgrade
	Upgrade-SPSite $path -VersionUpgrade -QueueOnly
 
	$loop = $true
 
	#loop while waiting for timer job to finish the site upgrade
	do{
		#pause for 10 seconds  to check status of site upgrade
		Start-Sleep -s 10
		Write-host "Waiting for SharePoint to begin the upgrade process for"$path  -foregroundcolor black -backgroundcolor green
 
		#get status of upgrade
		$upgradeSessionInfo = Get-SPSiteUpgradeSessionInfo -Site $path
        Write-host " Status: " $upgradeSessionInfo.Status
 
		switch($upgradeSessionInfo.Status){
			"Completed" 
			{      
				Write-host "Site upgrade complete for"$path  -foregroundcolor black -backgroundcolor green
        
				#enable branding features
				Write-host "Activating Branding feature on"$path  -foregroundcolor black -backgroundcolor green
				
				try{Enable-spfeature -Identity 12d06415-bc47-47ee-9f60-0f9a340b4398 -url $path -ErrorAction Stop}
				catch [System.Management.Automation.ActionPreferenceStopException]
				{
					if( !($_.Exception -is [System.Data.DuplicateNameException]) )
					{
						#unexpected error, bubble up
						throw    
					}
					else
					{
						#do nothing,  "feature is already activated at scope" exception
					}
				}

				#try
				#{
				#	Enable-spfeature -Identity 8ea2a086-984d-426a-b934-2808b7bf447d -url $path -ErrorAction Stop
				#}
				#	catch [System.Management.Automation.ActionPreferenceStopException]
				#{
				#
				#	if( !($_.Exception -is [System.Data.DuplicateNameException]) )
				#	{
				#		#unexpected error, bubble up
				#		throw    
				#	}
				#	else
				#	{
				#		#do nothing,  "feature is already activated at scope" exception
				#	}
				#}
 
 
 
			#update master page for each spweb in site collection
			Write-host "Updating master pages for"$path  -foregroundcolor black -backgroundcolor green
			UpdateMasterPage($path)
 
			Set-SPSite -Identity $path -SecondaryOwnerAlias "kelsey-seybold\altayl02" -OwnerAlias "kelsey-seybold\cmgrig01"
			Write-host "Updating site collection administrators for"$path  -foregroundcolor black -backgroundcolor green
 
			#upgrade complete
			$loop=$false
		}
 
		"Failed" 
		{
			#upgrade failed, move on to the next site
			Write-host "Site upgrade failed for"$path  -foregroundcolor red -backgroundcolor yellow;
			$loop=$false
		}
 
		"In Progress" {Write-host "Site upgrade in progress for"$path  -foregroundcolor black -backgroundcolor green}
		}
	}
	while($loop)
}
 
Function UpdateMasterPage($path){
	$site = Get-SPSite $path
	foreach ($web in $site.AllWebs) {
		$masterPath = $web.ServerRelativeUrl +"/_catalogs/masterpage/Upgrade2013.master";
		$web.CustomMasterUrl =$masterPath; 
		$web.MasterUrl =$masterPath; 
		$web.Update(); 
		Write-host "Upgrade2013.master has been applied to"$web.Url  -foregroundcolor black -backgroundcolor green
	}
	$site.Dispose()
}

[xml]$s = get-content UpgradeSitesINT.xml
Write-Host " "
Write-Host " "

foreach ($SiteCollection in $s.Upgrade.SiteCollections.SiteCollection)
{
	$SCName = $SiteCollection.getAttribute("Name")
	Echo $SCName
	$WebApp = $SiteCollection.getAttribute("HostHeaderWebApplication")
	Echo $WebApp
	$SiteURL = $SiteCollection.getAttribute("URL")
	Echo $SiteURL

	$FullSCURLPath = $WebApp + $SiteURL
	Echo "Name: $SCName -URL $FullSCURLPath"
	UpgradeSite($FullSCURLPath)
	
}

 



