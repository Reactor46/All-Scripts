#load hashtable of localized string
Import-LocalizedData -BindingVariable ConnectFunctions_LocalizedStrings -FileName ConnectFunctions.strings.psd1

## Default and minimum timeout value for sessions  = 3 minutes
## Maximum timeout value for sessions              = 15 minutes
##
$sessionOptionsTimeout = 180000;
if (($env:MsExchEmsTimeout -ne $null) -and ($env:MsExchEmsTimeout -gt 180000) -and ($env:MsExchEmsTimeout -lt 900001))
{
    $sessionOptionsTimeout = $env:MsExchEmsTimeout;
}
##

# help file specific for the function Connect-ExchangeServer.
# can be viewed by running: get-help Connect-ExchangeServer
function Connect-ExchangeServer ($ServerFqdn, [switch]$Auto, [switch]$Prompt, $UserName, $Forest,[switch]$ClearCache, $ClientApplication=$null, [switch]$AllowClobber)
{
#.EXTERNALHELP Connect-ExchangeServer-help.xml
	set-variable VerbosePreference -value Continue
	:connectScope do
	{
		if (!$Auto -and ($ServerFqdn -eq $null) -and !$Prompt)
		{
			_PrintUsageAndQuit
		}

		$useWIA = $true
		if (!($userName -eq $null))
		{
		    $credential = get-credential $username
		    $useWIA = $false
		}


		if (!($ServerFqdn -eq $null))
		{
			if ($Auto -or !($Forest -eq $null)) { _PrintUsageAndQuit }

			_OpenExchangeRunspace $ServerFqdn $credential $useWIA -ClientApplication:$ClientApplication
		}

		if ($Auto)
		{
			# We should provide the $credential before $Forest, and we cannot assume useWIA $true here. It should be read from $useWIA
			_AutoDiscoverAndConnect $credential $Forest -useWIA:$useWIA -ClientApplication:$ClientApplication
		}
		else
		{
			if (!($Forest -eq $null)) { _PrintUsageAndQuit }
		}

		Write-Host $ConnectFunctions_LocalizedStrings.res_0000
		$fqdn=read-host -prompt $ConnectFunctions_LocalizedStrings.res_0001
		_OpenExchangeRunspace $fqdn $credential $useWIA -ClientApplication:$ClientApplication
	}
	while ($false) #connectScope
	
	if ($ClearCache)
	{
		if ($AllowClobber)
		{
			ImportPSSession -ClearCache:$true -AllowClobber
		}
		else
		{
			ImportPSSession -ClearCache:$true
		}
	}
	else
	{
		if ($AllowClobber)
		{
			ImportPSSession -ClearCache:$false -AllowClobber
		}
		else
		{
			ImportPSSession -ClearCache:$false
		}
	}
}

function Discover-ExchangeServer ([System.Management.Automation.PSCredential] $Credential, 
									$Forest, 
									[bool]$UseWIA=$false, 
									[bool]$SuppressError=$false, 
									[Microsoft.Exchange.Data.ServerVersion]$CurrentVersion=$null,
									$ClientApplication = "EMC",
									$AllowRedirection=$true)
{
	set-variable VerbosePreference -value Continue
	:connectScope do
	{
		_AutoDiscoverAndConnect $Credential $Forest $UseWIA $SuppressError $CurrentVersion $ClientApplication $AllowRedirection
	}
	while ($false) #connectScope

	if (!($global:remoteSession -eq $null))
	{
		$global:remoteSession.ComputerName
		remove-pssession $global:remoteSession
	}
	else
	{
		if ($SuppressError -and ($error.Count -ne 0))
		{
			# When we suppress the error message, we still want the last error
			Write-Error $error[0]
		}
	}
}

function _AutoDiscoverAndConnect ([System.Management.Automation.PSCredential]$Credential, 
									$Forest, 
									[bool]$UseWIA=$false, 
									[bool]$SuppressError=$false, 
									[Microsoft.Exchange.Data.ServerVersion]$CurrentVersion=$null,
									$ClientApplication=$null,
									$AllowRedirection=$false)
{
	if ($Forest -eq $null)
	{
		$fqdn = _GetHostFqdn $CurrentVersion
		if ($fqdn -ne $null)
		{
			_OpenExchangeRunspace $fqdn $credential $UseWIA $SuppressError $ClientApplication $AllowRedirection
		}

		$forestName = _GetLocalForest
		$siteList   = _GetSites
		# Find servers in sites and connect, assuming the default behavior (connect to current site, 
		# if failed then connect to adjacent sites, if failed then connect to any random site that has exchange server)
		foreach ($siteDN in $SiteList)
		{
			# E14: 184676 - MSIT: Tools only installation cannot find a remote PowerShell endpoint
			# when the Forest name of the current site is not available, then we need to pass null value for forest  
			if (($siteDN -ne "*") -and ($forestName -ne $null))
			{
				$servers = _GetExchangeServersInSite $siteDN "/$forestName"
			}
			else
			{
				$servers = _GetExchangeServersInSite $siteDN 
			}
			if (($siteDN -ne "*")-and ($servers -eq $null))
			{
				$siteName = $siteDN.ToString().SubString(3).Split(",")[0]
				Write-Warning ($ConnectFunctions_LocalizedStrings.res_0002 -f $siteName)
			}
			if (($siteDN -eq "*")-and ($servers -eq $null))
			{
			        Write-Error $ConnectFunctions_LocalizedStrings.res_0003
			}
			_ConnectToAnyServer $servers $credential $UseWIA $SuppressError $CurrentVersion $ClientApplication $AllowRedirection
		}
	}
	else
	{
		$servers = _GetExchangeServersInSite "*" "/$Forest"
		_ConnectToAnyServer $servers $credential $UseWIA $SuppressError $CurrentVersion $ClientApplication $AllowRedirection
	}
}
function _GetSites()
{
	$localSite=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()
	if ($localSite -eq $null)
	{
		return #no site - no auto discovery
	}
	$siteDN=$localSite.GetDirectoryEntry().DistinguishedName
	$siteList = New-Object System.Collections.ArrayList
	if ($siteDN -ne $null)
	{
		[void] $SiteList.Add($siteDN)
	}
	# DCR 231555: TAp DCR:Respect site-link costs with remote powershell.
	# Now go one level to find adjacent sites based on site link and add them in the site list
	if ($localSite.SiteLinks -ne $null)
	{
		foreach ($siteLink in $localSite.SiteLinks)
		{
			$siteDN = $null
			# block going backwords
			if (($siteLink.Sites[0] -ne $null) -and ($siteLink.Sites[0].Name -ne $localSite.Name))
			{
				$siteDN = $siteLink.Sites[0].GetDirectoryEntry().DistinguishedName
			}
			elseif ($siteLink.Sites[1] -ne $null)
			{
				$siteDN = $siteLink.Sites[1].GetDirectoryEntry().DistinguishedName
			}
			if ($siteDN -ne $null)
			{ 
				[void] $SiteList.Add($siteDN)
			}
		}
	}
	# When no exchange server will be found in the current site or adjacent sites, then we should search in all sites
	[void] $SiteList.Add("*")
	$siteList
}	

function _GetLocalForest()
{
	[System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Forest.Name
}	

function _GetHostFqdn([Microsoft.Exchange.Data.ServerVersion]$CurrentVersion=$null)
{
	if (@(get-item HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\*role -erroraction:silentlycontinue).length -gt 0)
	{
		$setupRegistryEntry = get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -erroraction:silentlycontinue
		if ($CurrentVersion -ne $null -and $setupRegistryEntry -ne $null -and
			($CurrentVersion.Major -ne $setupRegistryEntry.MsiProductMajor -or
			$CurrentVersion.Minor -ne $setupRegistryEntry.MsiProductMinor -or
			$CurrentVersion.Build -ne $setupRegistryEntry.MsiBuildMajor -or
			$CurrentVersion.Revision -ne $setupRegistryEntry.MsiBuildMinor))
		{
			return $null
		}
		try
		{
			return [System.Net.Dns]::GetHostByName("LocalHost").HostName
		}
		catch
		{
			Write-Verbose $ConnectFunctions_LocalizedStrings.res_0004
			return $null
		}
	}
	else
	{
		return $null
	}
}

function _GetServerFqdnFromNetworkAddress($server)
{
   $server.properties["networkaddress"] |
      where {$_.ToString().StartsWith("ncacn_ip_tcp")} | %{$_.ToString().SubString(13)}
}


function _GetExchangeServersInSite($siteDN, $Forest=$null)
{
	$configNC=([ADSI]"LDAP:/$Forest/RootDse").configurationNamingContext
	$search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP:/$Forest/$configNC")
	$search.Filter = "(&(objectClass=msExchExchangeServer)(versionNumber>=1937801568)(msExchServerSite=$siteDN))"
	$search.PageSize=1000
	$search.PropertiesToLoad.Clear()
	[void] $search.PropertiesToLoad.Add("msexchcurrentserverroles")
	[void] $search.PropertiesToLoad.Add("networkaddress")
	[void] $search.PropertiesToLoad.Add("serialnumber")
	$search.FindAll()
}

function _GetCurrentVersionServers($servers, [Microsoft.Exchange.Data.ServerVersion]$CurrentVersion)
{
	$sameVersionServers=@()
	$sameBuildServers=@()
	$sameProductMajorMinorServers=@()
	foreach ($server in $servers)
	{
		[Microsoft.Exchange.Data.ServerVersion]$version=$null
		if([Microsoft.Exchange.Data.ServerVersion]::TryParseFromSerialNumber($server.Properties["serialnumber"][0], [ref]$version))
		{
			if($version -ne $null)
			{
				if ($version.Equals($CurrentVersion))
				{
					$sameVersionServers += $server
				}
				elseif (($version.Major -eq $CurrentVersion.Major) -and ($version.Minor -eq $CurrentVersion.Minor) -and ($version.Build -eq $CurrentVersion.Build))
				{
					$sameBuildServers += $server
				}
			}
		}
	}
	return @($sameVersionServers + $sameBuildServers)
}

function _GetWebServiceServers($servers, [Microsoft.Exchange.Data.ServerVersion]$CurrentVersion=$null)
{
	$allFfoWs = $servers | where {(($_.properties["msexchcurrentserverroles"][0] -band 65536) -and ($_.path -like "*/CN=*[WM][SF]0*"))}
	if ($allFfoWs -eq $null)
	{
		return @()
	}

	$webServiceServers = @()
	if ($CurrentVersion -eq $null)
	{
		$webServiceServers = $allFfoWs
	}
	else
	{
		$webServiceServers = @(_GetCurrentVersionServers $allFfoWs $CurrentVersion)
	}
    
    return $webServiceServers;	
}

function _GetCASServers($servers, [Microsoft.Exchange.Data.ServerVersion]$CurrentVersion=$null)
{
	$allCas = $servers | where {($_.properties["msexchcurrentserverroles"][0] -band 4)}
	if ($allCas -eq $null)
	{
		return @()
	}
	if ($CurrentVersion -eq $null)
	{
		return $allCas
	}
	else
	{
		return @(_GetCurrentVersionServers $allCas $CurrentVersion)
	}
}

function _GetCAFEServers($servers, [Microsoft.Exchange.Data.ServerVersion]$CurrentVersion=$null)
{
	$allCafes = $servers | where {($_.properties["msexchcurrentserverroles"][0] -band 1)}
	if ($allCafes -eq $null)
	{
		return @()
	}
	if ($CurrentVersion -eq $null)
	{
		return $allCafes
	}
	else
	{
		return @(_GetCurrentVersionServers $allCafes $CurrentVersion)
	}
}

# This function returns all non-cas and non-cafe servers, that includes mailbox, hub transport and um servers only
function _GetHubMailboxUMServers($servers, [Microsoft.Exchange.Data.ServerVersion]$CurrentVersion=$null)
{
	$allNonCasCafe = $servers | where {(($_.properties["msexchcurrentserverroles"][0] -band 4) -eq 0) -and (($_.properties["msexchcurrentserverroles"][0] -band 1) -eq 0) -and
					($_.properties["msexchcurrentserverroles"][0] -band 50)}
	if ($allNonCasCafe -eq $null)
	{
		return @()
	}
	if ($CurrentVersion -eq $null)
	{
		return $allNonCasCafe
	}
	else
	{
		return @(_GetCurrentVersionServers $allNonCasCafe $CurrentVersion)
	}
}

function _NewExchangeRunspace(
				[String]$fqdn, 
				[System.Management.Automation.PSCredential] $credential=$null, 
				[bool]$UseWIA=$true, 
				[bool]$SuppressError=$false,
				$ClientApplication=$null,
				$AllowRedirection=$false,
				[String]$targetServerUriParameter=$null)
{
	$hostFQDN = _GetHostFqdn
	if (($fqdn -ne $null) -and ($hostFQDN -ne $null) -and ($hostFQDN.ToLower() -eq $fqdn.ToLower()))
	{
	    $ServicesRunning = _CheckServicesStarted
	    if ($ServicesRunning -eq $false)
	    {
	        return
	    }
	}
	Write-Verbose ($ConnectFunctions_LocalizedStrings.res_0005 -f $fqdn)
	$so = New-PSSessionOption -OperationTimeout $sessionOptionsTimeout -IdleTimeout $sessionOptionsTimeout -OpenTimeout $sessionOptionsTimeout;
	$setupRegistryEntry = get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -erroraction:silentlycontinue
	if ( $setupRegistryEntry -ne $null)
	{
		$clientVersion = "{0}.{1}.{2}.{3}" -f $setupRegistryEntry.MsiProductMajor, $setupRegistryEntry.MsiProductMinor, $setupRegistryEntry.MsiBuildMajor, $setupRegistryEntry.MsiBuildMinor
		$connectionUri = "http://$fqdn/powershell?serializationLevel=Full;ExchClientVer=$clientVersion"
	}
	else
	{
		$connectionUri = "http://$fqdn/powershell?serializationLevel=Full"
	}
	
	if ($ClientApplication -ne $null)
	{
		$connectionUri = $connectionUri + ";clientApplication=$ClientApplication"
	}

	if ($targetServerUriParameter -ne $null)
	{
		$connectionUri = $connectionUri + ";TargetServer=$targetServerUriParameter"
	}
	
	$contents = 'New-PSSession -ConnectionURI "$connectionUri" -ConfigurationName Microsoft.Exchange -SessionOption $so'
	
	if (-not $UseWIA)
	{
		$contents = $contents + ' -Authentication Kerberos -Credential $credential'
	}
	if ($SuppressError)
	{
		$contents = $contents + ' -erroraction silentlycontinue'
	}
	if ($AllowRedirection)
	{
		$contents = $contents + ' -AllowRedirection'
	}
	[ScriptBlock] $command = $executioncontext.InvokeCommand.NewScriptBlock([string]::join("`n", $contents))
	$session=invoke-command -Scriptblock $command
	
  if (!$?)
	{
	  # ERROR_ACCESS_DENIED = 5
	  # ERROR_LOGON_FAILURE = 1326
	  if (!(5 -eq $error[0].exception.errorcode) -and
		  !(1326 -eq $error[0].exception.errorcode))
	  {
			#Write-Verbose ($ConnectFunctions_LocalizedStrings.res_0006 -f $fqdn)
			return
	  }
	  else
	  {
	    # no retries if we get 5 (access denied) or 1326 (logon failure)
	    #$REVIEW$ connectedFqdn is not set. Is it okay?
	    break connectScope
	  }
	}
  $session
}

function _OpenExchangeRunspace([String]$fqdn, 
				[System.Management.Automation.PSCredential] $credential, 
				[bool]$UseWIA=$true, 
				[bool]$SuppressError=$false,
				$ClientApplication=$null,
				$AllowRedirection=$false)
{
  $global:remoteSession = _NewExchangeRunspace $fqdn $credential $UseWIA $SuppressError $ClientApplication $AllowRedirection

  if ($global:remoteSession -ne $null)
  {
	  $global:connectedFqdn = $fqdn
	  break connectScope
  }
}

function CreateOrGetExchangeSession(
			[String]$fqdn, 
			[System.Management.Automation.PSCredential] $credential=$null, 
			[bool]$UseWIA=$true, 
			[bool]$SuppressError=$false,
			[String]$targetServerUriParameter=$null)
{
   $existingSession = @(Get-PSSession | Where {$_.ComputerName -eq $fqdn -and $_.Availability -eq 'Available' -and $_.ConfigurationName -eq 'Microsoft.Exchange'})

   if (($existingSession.Count -gt 0) -and ($targetServerUriParameter -eq $null))
   {
    if ($existingSession.Count -gt 1)
    {
      Write-Warning ($ConnectFunctions_LocalizedStrings.res_0007 -f $fqdn)
    }
    Write-Verbose ($ConnectFunctions_LocalizedStrings.res_0008 -f $fqdn)
    return $existingSession[0]
   }
   else
   {
    Write-Verbose ($ConnectFunctions_LocalizedStrings.res_0009 -f $fqdn)
    _NewExchangeRunspace $fqdn $credential $UseWIA $SuppressError $null $false $targetServerUriParameter
   }
}

function _ConnectToAnyServer ($servers, 
				$credential, 
				[bool]$UseWIA=$false, 
				[bool]$SuppressError=$false, 
				[Microsoft.Exchange.Data.ServerVersion]$CurrentVersion=$null,
				$ClientApplication=$null,
				$AllowRedirection=$false)
{
	if (($servers -eq $null) -or ($servers.Length -eq 0))
	{
		return
	}

	$cafe=@(_GetCAFEServers $servers $CurrentVersion)

	for($i=0;$i -lt $cafe.Length;$i++)
	{
		$fqdn = _GetServerFqdnFromNetworkAddress $cafe[($i+$start) % $cafe.Length]
		_OpenExchangeRunspace $fqdn $credential $UseWIA $SuppressError $ClientApplication $AllowRedirection
	}

	$webServiceServers=@(_GetWebServiceServers $servers $CurrentVersion)

	for($i=0;$i -lt $webServiceServers.Length;$i++)
	{
		$fqdn = _GetServerFqdnFromNetworkAddress $webServiceServers[($i+$start) % $webServiceServers.Length]
		_OpenExchangeRunspace $fqdn $credential $UseWIA $SuppressError $ClientApplication $AllowRedirection
	}

	$cas=@(_GetCASServers $servers)

	for($i=0;$i -lt $cas.Length;$i++)
	{
		$fqdn = _GetServerFqdnFromNetworkAddress $cas[($i+$start) % $cas.Length]
		_OpenExchangeRunspace $fqdn $credential $UseWIA $SuppressError $ClientApplication $AllowRedirection
	}

	$other=@(_GetHubMailboxUMServers $servers $CurrentVersion)

	for($i=0;$i -lt $other.Length;$i++)
	{
		$fqdn = _GetServerFqdnFromNetworkAddress $other[($i+$start) % $other.Length]
		_OpenExchangeRunspace $fqdn $credential $UseWIA $SuppressError $ClientApplication $AllowRedirection
	}
}

function _PrintUsageAndQuit
{
  $ConnectFunctions_LocalizedStrings.res_0012
  break connectScope
}

function _CheckServicesStarted
{
    #List of services that should be running while connecting to remote server
    $DependentServices = @('w3svc')
    foreach ($Service in $DependentServices)
    {
        $ServiceState = Get-Service $Service -ea SilentlyContinue
        $ServiceNameFull = $ServiceState.Name + ' (' + $ServiceState.DisplayName + ')'
        if ($ServiceState.Status -ne 'Running')
        {
            Write-Warning ($ConnectFunctions_LocalizedStrings.res_0010 -f $ServiceNameFull)
            return $false
        }
    }
    return $true
}


# This function is used to discover ECP when EMC connects to localOnPremise
function Discover-EcpVirtualDirectoryForEmc ([Microsoft.Exchange.Data.ServerVersion]$CurrentVersion, [bool]$UseWIA=$false, $TargetServerFqdn=$null)
{
	# get the siteDN, forest
	$forestName = _GetLocalForest
	$siteList   = _GetSites

	# Get the localhost's fqdn so we can check the vdir on local machine first
	$hostFqdn = _GetHostFqdn

	# declare $version to store the return value from TryParseFromSerialNumber
	[Microsoft.Exchange.Data.ServerVersion]$version=$null

	# declare the DirectorySearcher of ECP vdir for later use
	$search = new-object DirectoryServices.DirectorySearcher
	$search.Filter = "(&(objectClass=msExchEcpVirtualDirectory))"
	$search.PageSize=1000
	$search.PropertiesToLoad.Clear()
	[void] $search.PropertiesToLoad.Add("msexchinternalauthenticationmethods")
	[void] $search.PropertiesToLoad.Add("msexchinternalhostname")
	[void] $search.PropertiesToLoad.Add("msexchexternalhostname")


	# When EMC is connected as RemoteOnPremise, it's possible that the server is in the same forest as the host machine
	# so we check for that before we do autodiscover. This variable is used to do that
	$shouldDiscover=$false

	# first look at local site, then expand scope to forest
	foreach ($siteDN in $SiteList)
	{

		# Init variables
		if (($siteDN -ne "*") -and ($forestName -ne $null))
		{
			$allServers = _GetExchangeServersInSite $siteDN "/$forestName"
		}
		else
		{
			$allServers = _GetExchangeServersInSite $siteDN 
		}

		if($allServers -ne $null)
		{
		  # if $shouldDiscover is false, we should check whether we should do discovery or not first
		  if($shouldDiscover -eq $false){
			if($TargetServerFqdn -ne $null)
			{
			  # TargetServerFqdn is specified, which means we are connecting using remoteOnPremise
			  # so we need to check whether the targetServer is in the same forest or not.
			  foreach ($server in $allServers)
			  {
				$tempFqdn = _GetServerFqdnFromNetworkAddress $server
				if (($TargetServerFqdn -eq $tempFqdn))
				{
				  $shouldDiscover=$true
				  break
				}
			  }
			}else{
			  $shouldDiscover=$true
			}

			if($shouldDiscover -eq $false){
			  continue
			}
		  }

		  $servers = _GetCASServers $allServers

		  #continue if there's no CAS
		  if(($servers -eq $null) -or ($servers.length -eq 0)){
			continue
		  }

		  # group the servers by version
		  $localServer=$null
		  $sameVersionServers=@()
		  $newerVersionServers=@()
		  $olderServers=@()
		  foreach ($server in $servers) {
			# if the version can't be parsed, we will skip this server
			if([Microsoft.Exchange.Data.ServerVersion]::TryParseFromSerialNumber($server.Properties["serialnumber"][0], [ref]$version))
			{
			  # only look for server with same major version
			  if($version.Major -eq $CurrentVersion.Major)
			  {
				if($version.Equals($CurrentVersion))
				{
				  $fqdn = _GetServerFqdnFromNetworkAddress $server
				  if (($hostFQDN -ne $null) -and ($hostFQDN -eq $fqdn)){
					# if it's localhost, insert it into the front of the array so it's checked before others
					$sameVersionServers = @($server)+$sameVersionServers
				  }else{
					$sameVersionServers += $server
				  }
				}elseif($version.ToInt() -gt $CurrentVersion.ToInt()){
				  $newerVersionServers += $server
				}else{
				  $olderServers += $server
				}
			  }
			}
		  }

		  # find a suitable vdir, the order here shouldn't be changed
		  foreach ($servers in @($sameVersionServers,$newerVersionServers,$olderServers))
		  {
			$fbaVdirs=@()
			$allVdirs=@()
			foreach($server in $servers)
			{
			  Write-Verbose ($ConnectFunctions_LocalizedStrings.res_0011 -f $server)
			  $search.SearchRoot=[ADSI]$server.Path
			  $vdirs=$search.FindAll();
			  $tempFba=$null
			  if($vdirs -ne $null)
			  {
				$selected,$tempFba= _SelectVdir $vdirs $useWIA
				if( $selected -ne $null )
				{
				  return _GetUrl $selected
				}else{
				  $fbaVdirs+=$tempFba
				  $allVdirs+=$vdirs
				}
			  }
			}

			if($fbaVdirs.length -gt 0)
			{
			  # we can't find a WIA vdir, so just return a fbaVdir
			  return _GetUrl $fbaVdirs[0]
			}
			elseif($allVdirs.length -gt 0)
			{
			  # we can't even find a FBA vdir, so just return any vdir
			  return _GetUrl $vdirs[0]
			}
		  }
		}
	}
}

# select vdir based on the criteria
# Parameters
#	vdirs: a list of vdirs to be selected from
#	useWIA: True->select a WIA; False->select a FBA
# Return:
#	retVal: vdir that matches the authentication mode requested
#	fbaVdirs: return the list of fbaVdirs for caller
function _SelectVdir ($vdirs, [bool]$useWIA)
{
	$fbaVdirs=@()
	$retVal=$null
	
	# randomize the start position. There are two benefits
	# 1. If the first server is down, this randomization will allow us to pick other servers
	# 2. It helps load ballancing.
	$random = New-Object System.Random
	$start = $random.Next(0, $vdirs.Count)
	
	for($i=$start; ($i-$start) -lt $vdirs.Count; $i++)
	{
		$vdir = $vdirs[$i % $vdirs.Count]

		# search for WIA vdir
		if($useWIA)
		{
			if(_IsWIA $vdir)
			{
				$retVal=$vdir
				break
			}
		}

		# see if it's FBA
		if(_IsFBA $vdir)
		{
			if($useWIA)
			{
				# save the vdirs in case we can't find a WIA vdir later
				$fbaVdirs += $vdir
			}
			else
			{
				# we are looking for FBA here, so just return
				$retVal = $vdir
				break
			}
		}
	}

	return $retVal, $fbaVdirs
}

# determine if authentication mode is WIA. According to AuthenticationMethodFlags, Ntlm is 2, WindowsIntegrated is 16.
# Exchange's task will normally set them both. But actually normally windows supports both, so I'll return true if either
# of them is set.Also note that we are checking the internal autentication flag.
function _IsWIA ($vdir)
{
	_InternalAuthenticationMatchesAny $vdir 18
}

# determine if authentication mode is FBA. According to AuthenticationMethodFlags, Fba is 4.
# Note that Exchange's task will set Basic to true when FBA is true.
# Note that I'm not checking external authentication flags. Task Set-EcpVdir -FormAuthentication:$false, it's only 
# updating the internal flag, 
function _IsFBA ($vdir)
{
	_InternalAuthenticationMatchesAny $vdir 4
}

# select vdir based on the criteria
# Parameters
#	vdir: vdir to check
#	flag: target flag
# Return:
#	true if vdir's msexchinternalauthenticationmethods has any of the $flag
function _InternalAuthenticationMatchesAny($vdir, $flag)
{
	if($vdir.Properties["msexchinternalauthenticationmethods"] -ne $null)
	{
		($vdir.Properties["msexchinternalauthenticationmethods"][0] -band $flag) -ne 0
	}
	else
	{
		return $false
	}
}

# get the url from a vdir
function _GetUrl ($vdir)
{
	if($vdir.Properties["msexchinternalhostname"] -ne $null)
	{
		$vdir.Properties["msexchinternalhostname"]
	}elseif($vdir.Properties["msexchexternalhostname"] -ne $null)
	{
		$vdir.Properties["msexchexternalhostname"]
	}
}

# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUzjrOdtVijCqNaHpaci5YcEMX
# wWOgghhkMIIEwzCCA6ugAwIBAgITMwAAAKxjFufjRlWzHAAAAAAArDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwNTAzMTcxMzIz
# WhcNMTcwODAzMTcxMzIzWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkMwRjQtMzA4Ni1ERUY4MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnyHdhNxySctX
# +G+LSGICEA1/VhPVm19x14FGBQCUqQ1ATOa8zP1ZGmU6JOUj8QLHm4SAwlKvosGL
# 8o03VcpCNsN+015jMXbhhP7wMTZpADTl5Ew876dSqgKRxEtuaHj4sJu3W1fhJ9Yq
# mwep+Vz5+jcUQV2IZLBw41mmWMaGLahpaLbul+XOZ7wi2+qfTrPVYpB3vhVMwapL
# EkM32hsOUfl+oZvuAfRwPBFxY/Gm0nZcTbB12jSr8QrBF7yf1e/3KSiqleci3GbS
# ZT896LOcr7bfm5nNX8fEWow6WZWBrI6LKPx9t3cey4tz0pAddX2N6LASt3Q0Hg7N
# /zsgOYvrlwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCFXLAHtg1Boad3BTWmrjatP
# lDdiMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAEY2iloCmeBNdm4IPV1pQi7f4EsNmotUMen5D8Dg4rOLE9Jk
# d0lNOL5chmWK+d9BLG5SqsP0R/gqph4hHFZM4LVHUrSxQcQLWBEifrM2BeN0G6Yp
# RiGB7nnQqq86+NwX91pLhJ5LBzJo+EucWFKFmEBXLMBL85fyCusCk0RowdHpqh5s
# 3zhkMgjFX+cXWzJXULfGfEPvCXDKIgxsc5kUalYie/mkCKbpWXEW6gN+FNPKTbvj
# HcCxtcf9mVeqlA5joTFe+JbMygtOTeX0Mlf4rTvCrf3kA0zsRJL/y5JdihdxSP8n
# KX5H0Q2CWmDDY+xvbx9tLeqs/bETpaMz7K//Af4wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBK4wggSqAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBwjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUwkDNZMzLH0tRyLxKRbaKgRdpYJkwYgYKKwYB
# BAGCNwIBDDFUMFKgKoAoAEMAbwBuAG4AZQBjAHQARgB1AG4AYwB0AGkAbwBuAHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBADiumoY+nXcWYMb2Fl78VWPEHp4OgdCjWehziQwo+qwN
# 7uZISHodjO9WnLrnOUXP5PU1FF2N6sQz/gbwSFQT/jSmy1c5MQh3EiIhl2kfBzFq
# gdtv4jEbc5FVASGCGnHxJGZteOVKEeRt4HamqQBysG6ZwcAymXF8TgnVDRVi0Xjy
# PJIMugOvfQ3FE70KH3pxiTVf/n5uGm8Qj0UwyZ5Zaz1Z+Vj1scYz+OP1AycYD0vD
# c5ZURBwzs03MoF0aeyNQ4yXXw7pgKVM0lRqlUUwnjIUwTWJhUN7bYsYegIr/NL1Z
# P+yiiDu5ZK/9nDUWKWWuNVet2hF9c8iGYIo174+JL72hggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAA
# rGMW5+NGVbMcAAAAAACsMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ0MzVaMCMGCSqGSIb3DQEJ
# BDEWBBReX5GfVDDA1c/gmPvA3NmCjtwYNDANBgkqhkiG9w0BAQUFAASCAQBvACax
# mNej95g0nCW91St8Oe+wY8KDrpFccSZUre+tVTwb/PAklrO2rtLL33/hCmySVVbm
# sQml4I59Ndpp4p4E+MHGWx/uVoe3+2+JO3nL5Enqw157DVSBY3eFEkCvSLPicZlO
# q03dwHFZQcTJNR4GX++fTyM+8CGVISOwjuT9vbuZxJ03nxQfgWtoWiCLgTAKT7yM
# 353Q0eBvyZVIxAUKOnrZbIQPPd0/26Jc5+rgNyKC90vahiT3s0n3fbcijwtoRISM
# d9Dw7e9YkwaJ3+9jEvbpiRUvGYvsqeR6LOMbeO0jzTYFIr5rOKi8WcdigEdHi7A+
# yDEUftn4DdTCWOE2
# SIG # End signature block
