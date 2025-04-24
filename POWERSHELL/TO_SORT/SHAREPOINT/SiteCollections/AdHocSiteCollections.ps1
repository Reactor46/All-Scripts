#Add Sites/SubSites

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction "SilentlyContinue" | Out-Null
        
    $ErrorActionPreference = "Stop"
    
    if ($Verbose -eq $null) { $Verbose = $false }

# Ensure the SPTimerService is started on each Application Server
foreach ($server in (get-spserver | Where {$_.Role -eq "Application"}) )
{
	Write-Host "Starting SPTimerService on each Application Server"
	$server.Name
	$service = Get-WmiObject -computer $server.Name Win32_Service -Filter "Name='SPTimerV4'"
	$service.InvokeMethod('StopService',$Null)
	start-sleep -s 5
	$service.InvokeMethod('StartService',$Null)
	start-sleep -s 5
	$service.State
}

[xml]$s = get-content AdHocSiteCollections.xml
Write-Host " "
Write-Host " "

foreach ($SiteCollection in $s.Setup.SiteCollections.SiteCollection)
{
	$SCName = $SiteCollection.getAttribute("Name")
	$CTDB = $SiteCollection.getAttribute("ContentDatabase")
	$WebApp = $SiteCollection.getAttribute("HostHeaderWebApplication")
	$SiteURL = $SiteCollection.getAttribute("Url")
	$OA = $SiteCollection.getAttribute("OwnerAlias")
	$SOA = $SiteCollection.getAttribute("SecondaryOwnerAlias")
	$Desc = $SiteCollection.getAttribute("Description")
	$Template = $SiteCollection.getAttribute("Template")
	$Language = $SiteCollection.getAttribute("Language")

	$FullSCURLPath = $WebApp + $SiteURL
	Echo "-URL $FullSCURLPath -OwnerAlias $OA ContentDB: $CTDB Name: $SCName Description: $Desc Language: $Language"

	echo "Checking for Site Collection: $FullSCURLPath"
	$TestSC = Get-SPSite -Identity $FullSCURLPath -ErrorVariable err -ErrorAction SilentlyContinue 
	if ($err) 
	{ 
	   	echo "Checking for Content Database: $CTDB"

		$TestDB =get-spcontentdatabase $CTDB -ErrorVariable err -ErrorAction SilentlyContinue
		if ($err)
		{
			echo "Creating Database: $CTDB"
			new-spcontentdatabase $CTDB -WebApplication $WebApp
		}
		else
		{
			echo "$TestDB already exists"

		}
		echo "Creating Site Collection: $SCName"
		echo "Site CollectionURL: " $FullSCURLPath
		new-spsite -URL $FullSCURLPath -OwnerAlias $OA -SecondaryOwnerAlias $SOA -ContentDatabase $CTDB -Name $SCName -Description $Desc -Language $Language -Template $Template
		echo "Site Collection Created"
		Write-Host "  "
		Write-Host "************************************************************************************************"
		Write-Host " "



	} 
	else 
	{
		echo "Site: $FullSCURLPath already exists"

	}
	#Activate Site Features

	foreach($SF in $SiteCollection.Feature)
	{
		Write-Host " "
		if($SF -ne $Null)
		{
			$SFName = $SF.getAttribute("Name")
			$SFGuid = $SF.getAttribute("Guid")
			

			$TestSF = get-spfeature -Site $FullSCURLPath -Identity $SFGuid -ErrorVariable sferr -ErrorAction SilentlyContinue
			if ($sferr)
			{
				echo "Activating Site Feature: $SFName"
				enable-spfeature -URL $FullSCURLPath -identity $SFGuid -force
			}
			else
			{
				echo "$SFName is needs to installed and activated manually"
			}
		}
	}


	#Create SubSites
	foreach($SS in $SiteCollection.SubSite)
	{ 
	Write-Host " "
		if($SS -ne $Null)
		{
			$SSName = $SS.getAttribute("Name")
			$SSURL = $SS.getAttribute("Url")
			$SSDescription = $SS.getAttribute("Description")
			$SSTemplate = $SS.getAttribute("Template")
			$SSLanguage = $SS.getAttribute("Language")

			$NewURL = $FullSCURLPath + $SSURL
			echo "SPWEb: $NewURL -template $SSTemplate -name $SSName -Description $SSDescription -Language $SSLanguage" -AddToQuickLaunch -UniquePermissions -UseParentTopNav
			$TestSS =get-spweb $NewURL -ErrorVariable sserr -ErrorAction SilentlyContinue
			if ($sserr)
			{
				echo "Creating SPWeb: $NewURL"
				new-SPWeb $NewURL -template $SSTemplate -name $SSName -Description $SSDescription -Language $SSLanguage -AddToQuickLaunch -UniquePermissions -UseParentTopNav
			}
			else
			{
				echo "$NewURL already exists"

			}
			
			#Activate Web Features

			foreach($WF in $SS.Feature)
			{
				if($WF -ne $Null)
				{
					$WFName = $WF.getAttribute("Name")
					$WFGuid = $WF.getAttribute("Guid")
			

					$TestWF = get-spfeature -Site $NewURL -Identity $WFGuid -ErrorVariable WFerr -ErrorAction SilentlyContinue
					if ($WFerr)
					{
						echo "Activating Site Feature: $WFName"
						enable-spfeature -URL $NewURL -identity $WFGuid -force
					}
					else
					{
						echo "$WFName is already activated"
					}
				}
			}


			foreach($SubSS in $SS.SubSS)
			{
				Write-Host " "
				if($SubSS -ne $Null)
				{
					$SubSSName = $SubSS.getAttribute("Name")
					$SubSSURL = $SubSS.getAttribute("Url")
					$SubSSDescription = $SubSS.getAttribute("Description")
					$SubSSTemplate = $SubSS.getAttribute("Template")
					$SubSSLanguage = $SubSS.getAttribute("Language")

					$SubNewURL = $NewURL + $SubSSURL
					echo "SPWEb: $SubNewURL -template $SubSSTemplate -name $SubSSName -Description $SubSSDescription -Language $SubSSLanguage" -AddToQuickLaunch -UniquePermissions -UseParentTopNav
					$TestSubSS = get-spweb $SubNewURL -ErrorVariable SubSSerr -ErrorAction SilentlyContinue
					if ($SubSSerr)
					{
						echo "Creating SPWeb: $NewSubURL"
						new-SPWeb $SubNewURL -template $SubSSTemplate -name $SubSSName -Description $SubSSDescription -Language $SubSSLanguage -AddToQuickLaunch -UniquePermissions -UseParentTopNav
					}
					else
					{
						echo "$NewSubURL already exists"
					}

					#Activate SubWeb Features

					foreach($SubWF in $SUBSS.Feature)
					{
						if($SubWF -ne $Null)
						{
							$SubWFName = $SubWF.getAttribute("Name")
							$SubWFGuid = $SubWF.getAttribute("Guid")
			

							$TestSubWF = get-spfeature -Site $SubNewURL -Identity $SubWFGuid -ErrorVariable SubWFerr -ErrorAction SilentlyContinue
							if ($SubWFerr)
							{
								echo "Activating Site Feature: $SubWFName"
								enable-spfeature -URL $SubNewURL -identity $SubWFGuid -force
							}
							else
							{
								echo "$SubWFName is already activated"
							}
						}
					}


					foreach($SubSSS in $SubSS.SubSSS)
					{
						Write-Host " "
						if($SubSSS -ne $Null)
						{
							$SubSSSName = $SubSSS.getAttribute("Name")
							$SubSSSURL = $SubSSS.getAttribute("Url")
							$SubSSSDescription = $SubSSS.getAttribute("Description")
							$SubSSSTemplate = $SubSSS.getAttribute("Template")
							$SubSSSLanguage = $SubSSS.getAttribute("Language")

							$SubNewURL = $NewURL + $SubSSSURL
							echo "SPWeb: $SubNewURL -template $SubSSSTemplate -name $SubSSSName -Description $SubSSSDescription -Language $SubSSSLanguage" -AddToQuickLaunch -UniquePermissions -UseParentTopNav
							$TestSubSSS = get-spweb $SubNewURL -ErrorVariable SubSSSerr -ErrorAction SilentlyContinue
							if ($SubSSSerr)
							{
								echo "Creating SPWeb: $NewSubURL"
								new-SPWeb $SubNewURL -template $SubSSSTemplate -name $SubSSSName -Description $SubSSSDescription -Language $SubSSSLanguage -AddToQuickLaunch -UniquePermissions -UseParentTopNav
							}
							else
							{
								echo "$NewSubURL already exists"
							}

							#Activate SubWeb Features

							foreach($SubSWF in $SUBSSS.Feature)
							{
								if($SuSbWF -ne $Null)
								{
									$SubSWFName = $SubSWF.getAttribute("Name")
									$SubSWFGuid = $SubSWF.getAttribute("Guid")
			

									$TestSubSWF = get-spfeature -Site $SubNewURL -Identity $SubSWFGuid -ErrorVariable SubSWFerr -ErrorAction SilentlyContinue
									if ($SubSWFerr)
									{
										echo "Activating Site Feature: $SubSWFName"
										enable-spfeature -URL $SubNewURL -identity $SubSWFGuid -force
									}
									else
									{
										echo "$SubSWFName is already activated"
									}
								}
							}
						}
					}

		
				}
			}
		}
	}

}



