function Get-ExchangeInfo
{
	param($E2010,$ExchangeServer,$Mailboxes,$Databases,$Hybrids)
	
	# Set Basic Variables
	$MailboxCount = 0
	$RollupLevel = 0
	$RollupVersion = ""
    $ExtNames = @()
    $IntNames = @()
    $CASArrayName = ""
	
	# Get WMI Information
	$tWMI = Get-WmiObject Win32_OperatingSystem -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue
	if ($tWMI)
	{
		$OSVersion = $tWMI.Caption.Replace("(R)","").Replace("Microsoft ","").Replace("Enterprise","Ent").Replace("Standard","Std").Replace(" Edition","")
		$OSServicePack = $tWMI.CSDVersion
		$RealName = $tWMI.CSName.ToUpper()
	} else {
		Write-Warning "Cannot detect OS information via WMI for $($ExchangeServer.Name)"
		$OSVersion = "N/A"
		$OSServicePack = "N/A"
		$RealName = $ExchangeServer.Name.ToUpper()
	}
	$tWMI=Get-WmiObject -query "Select * from Win32_Volume" -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue
	if ($tWMI)
	{
		$Disks=$tWMI | Select Name,Capacity,FreeSpace | Sort-Object -Property Name
	} else {
		Write-Warning "Cannot detect OS information via WMI for $($ExchangeServer.Name)"
		$Disks=$null
	}
	
	# Get Exchange Version
	if ($ExchangeServer.AdminDisplayVersion.Major -eq 6)
	{
		$ExchangeMajorVersion = "$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
		$ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.FilePatchLevelDescription.Replace("Service Pack ","")
	} elseif ($ExchangeServer.AdminDisplayVersion.Major -eq 15 -and $ExchangeServer.AdminDisplayVersion.Minor -eq 1) {
        $ExchangeMajorVersion = [double]"$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
        $ExchangeSPLevel = 0
    } else {
		$ExchangeMajorVersion = $ExchangeServer.AdminDisplayVersion.Major
		$ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.Minor
	}
	# Exchange 2007+
	if ($ExchangeMajorVersion -ge 8)
	{
		# Get Roles
		$MailboxStatistics=$null
	    [array]$Roles = $ExchangeServer.ServerRole.ToString().Replace(" ","").Split(",");
        # Add Hybrid "Role" for report
        if ($Hybrids -contains $ExchangeServer.Name)
        {
            $Roles+="Hybrid"
        }
		if ($Roles -contains "Mailbox")
		{
			$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
			if ($ExchangeServer.Name.ToUpper() -ne $RealName)
			{
				$Roles = [array]($Roles | Where {$_ -ne "Mailbox"})
				$Roles += "ClusteredMailbox"
			}
			# Get Mailbox Statistics the normal way, return in a consitent format
			$MailboxStatistics = Get-MailboxStatistics -Server $ExchangeServer | Select DisplayName,@{Name="TotalItemSizeB";Expression={$_.TotalItemSize.Value.ToBytes()}},@{Name="TotalDeletedItemSizeB";Expression={$_.TotalDeletedItemSize.Value.ToBytes()}},Database
	    }
        # Get HTTPS Names (Exchange 2010 only due to time taken to retrieve data)
        if ($Roles -contains "ClientAccess" -and $E2010)
        {
            
            Get-OWAVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
            Get-WebServicesVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
            Get-OABVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
            Get-ActiveSyncVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
            if (Get-Command Get-MAPIVirtualDirectory -ErrorAction SilentlyContinue)
            {
                Get-MAPIVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
            }
            if (Get-Command Get-ClientAccessService -ErrorAction SilentlyContinue)
            {
                $IntNames+=(Get-ClientAccessService -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host
            } else {
                $IntNames+=(Get-ClientAccessServer -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host
            }
            
            if ($ExchangeMajorVersion -ge 14)
            {
                Get-ECPVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
            }
            $IntNames = $IntNames|Sort-Object -Unique
            $ExtNames = $ExtNames|Sort-Object -Unique
            $CASArray = Get-ClientAccessArray -Site $ExchangeServer.Site.Name
            if ($CASArray)
            {
                $CASArrayName = $CASArray.Fqdn
            }
        }

		# Rollup Level / Versions (Thanks to Bhargav Shukla http://bit.ly/msxGIJ)
		if ($ExchangeMajorVersion -ge 14) {
            $RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
        } else {
			$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\461C2B4266EDEF444B864AD6D9E5B613\\Patches"
		}
		$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);
		if ($RemoteRegistry)
		{
			$RUKeys = $RemoteRegistry.OpenSubKey($RegKey).GetSubKeyNames() | ForEach {"$RegKey\\$_"}
			if ($RUKeys)
			{
				[array]($RUKeys | %{$RemoteRegistry.OpenSubKey($_).getvalue("DisplayName")}) | %{
					if ($_ -like "Update Rollup *")
					{
						$tRU = $_.Split(" ")[2]
						if ($tRU -like "*-*") { $tRUV=$tRU.Split("-")[1]; $tRU=$tRU.Split("-")[0] } else { $tRUV="" }
						if ([int]$tRU -ge [int]$RollupLevel) { $RollupLevel=$tRU; $RollupVersion=$tRUV }
					}
				}
			}
        } else {
			Write-Warning "Cannot detect Rollup Version via Remote Registry for $($ExchangeServer.Name)"
		}
        # Exchange 2013 CU or SP Level
        if ($ExchangeMajorVersion -ge 15)
		{
			$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Microsoft Exchange v15"
		    $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);
		    if ($RemoteRegistry)
		    {
			    $ExchangeSPLevel = $RemoteRegistry.OpenSubKey($RegKey).getvalue("DisplayName")
                if ($ExchangeSPLevel -like "*Service Pack*" -or $ExchangeSPLevel -like "*Cumulative Update*")
                {
			        $ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2013 ","");
                    $ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2016 ","");
                    $ExchangeSPLevel = $ExchangeSPLevel.Replace("Service Pack ","SP");
                    $ExchangeSPLevel = $ExchangeSPLevel.Replace("Cumulative Update ","CU"); 
                } else {
                    $ExchangeSPLevel = 0;
                }
            } else {
			    Write-Warning "Cannot detect CU/SP via Remote Registry for $($ExchangeServer.Name)"
		    }
        }
		
	}
	# Exchange 2003
	if ($ExchangeMajorVersion -eq 6.5)
	{
		# Mailbox Count
		$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
		# Get Role via WMI
		$tWMI = Get-WMIObject Exchange_Server -Namespace "root\microsoftexchangev2" -Computername $ExchangeServer.Name -Filter "Name='$($ExchangeServer.Name)'"
		if ($tWMI)
		{
			if ($tWMI.IsFrontEndServer) { $Roles=@("FE") } else { $Roles=@("BE") }
		} else {
			Write-Warning "Cannot detect Front End/Back End Server information via WMI for $($ExchangeServer.Name)"
			$Roles+="Unknown"
		}
		# Get Mailbox Statistics using WMI, return in a consistent format
		$tWMI = Get-WMIObject -class Exchange_Mailbox -Namespace ROOT\MicrosoftExchangev2 -ComputerName $ExchangeServer.Name -Filter ("ServerName='$($ExchangeServer.Name)'")
		if ($tWMI)
		{
			$MailboxStatistics = $tWMI | Select @{Name="DisplayName";Expression={$_.MailboxDisplayName}},@{Name="TotalItemSizeB";Expression={$_.Size}},@{Name="TotalDeletedItemSizeB";Expression={$_.DeletedMessageSizeExtended }},@{Name="Database";Expression={((get-mailboxdatabase -Identity "$($_.ServerName)\$($_.StorageGroupName)\$($_.StoreName)").identity)}}
		} else {
			Write-Warning "Cannot retrieve Mailbox Statistics via WMI for $($ExchangeServer.Name)"
			$MailboxStatistics = $null
		}
	}	
	# Exchange 2000
	if ($ExchangeMajorVersion -eq "6.0")
	{
		# Mailbox Count
		$MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
		# Get Role via ADSI
		$tADSI=[ADSI]"LDAP://$($ExchangeServer.OriginatingServer)/$($ExchangeServer.DistinguishedName)"
		if ($tADSI)
		{
			if ($tADSI.ServerRole -eq 1) { $Roles=@("FE") } else { $Roles=@("BE") }
		} else {
			Write-Warning "Cannot detect Front End/Back End Server information via ADSI for $($ExchangeServer.Name)"
			$Roles+="Unknown"
		}
		$MailboxStatistics = $null
	}
	
	# Return Hashtable
	@{Name					= $ExchangeServer.Name.ToUpper()
	 RealName				= $RealName
	 ExchangeMajorVersion 	= $ExchangeMajorVersion
	 ExchangeSPLevel		= $ExchangeSPLevel
	 Edition				= $ExchangeServer.Edition
	 Mailboxes				= $MailboxCount
	 OSVersion				= $OSVersion;
	 OSServicePack			= $OSServicePack
	 Roles					= $Roles
	 RollupLevel			= $RollupLevel
	 RollupVersion			= $RollupVersion
	 Site					= $ExchangeServer.Site.Name
	 MailboxStatistics		= $MailboxStatistics
	 Disks					= $Disks
     IntNames				= $IntNames
     ExtNames				= $ExtNames
     CASArrayName			= $CASArrayName
	}	
}
