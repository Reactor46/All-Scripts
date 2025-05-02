# .SYNOPSIS
#   This is a wrapper script for bulk processing various MRS MoveRequest cmdlets.  
# 
# .DESCRIPTION
#   There are actually two independant scripts
#   First one: Executes "[New|Set|Resume|Remove]-MoveRequest" cmdlets by reading their parameter data from the passed CSV data. This will be used in XFR (cross forest relocation) work.
#	Second one: Injects move requests by reading the input from a given file
#  
# Parameters for the first script:
#
# .PARAMETER Records
#   Cmdlet records data. Each record contains the cmdlet name and parameters
#
# .PARAMETER TenantName
#   Name of the tenant that is used in checking if the offline GLS override is specified for this tenant. For protection purposes.
#
# Common parameters:
#
# .PARAMETER LogFile
#   Log file name. Logs are written to the local machine where the script is actually executing.
#	If not specified, it will be created in ExchangeInstallFolder\logging\XFR folder
#
# .PARAMETER Test
#   If set to "True", performs the simulation only.
#
# Parameters of the second script:
#
# .PARAMETER $BatchName
#   Batchname of the injected moves
#
# .PARAMETER $InputFile
#   The file that contains the move injection data
#
# .PARAMETER IgnoreExcludedFromInitialProvisioning
#   Ignores the IsExcludedFromInitialProvisioning db flag
#
#
# .EXAMPLE
#   
#   Given these sample CSV files:

#   Sample1:
#   Operation,Identity,RemoteHostName,RemoteOrganizationName,TargetDeliveryDomain,SkipMoving,InternalFlags,Remote,Protect,PreventCompletion,Confirm,CheckInitialProvisioningSetting,BadItemLimit,TargetDatabase,ArchiveTargetDatabase,SuspendWhenReadyToComplete,Priority,StartAfter,CompleteAfter,LargeItemLimit,AllowLargeItems,BatchName,DoNotPreserveMailboxSignature,AcceptLargeDataLoss,ForceOffline
#   new-moverequest,test4@kc2.extest.microsoft.com,EXHV-7909.EXHV-7909dom.extest.microsoft.com,kc2.extest.microsoft.com,kc2.extest.microsoft.com,,"ConvertSourceToMeu,ResolveServer,UseTcp,CrossResourceForest,SkipProvisioningCheck",1,1,,,1,444,,,1,High,,,,,,,,
#   new-moverequest,test2@kc2.extest.microsoft.com,EXHV-7909.EXHV-7909dom.extest.microsoft.com,kc2.extest.microsoft.com,kc2.extest.microsoft.com,,"ConvertSourceToMeu,ResolveServer,UseTcp,CrossResourceForest,SkipProvisioningCheck",1,1,1,,,,,,,,,,,,,,,
#   set-moverequest,test4@kc2.extest.microsoft.com,,,,,,,,0,0,,,,,,,,,,,,,,
#   resume-moverequest,test4@kc2.extest.microsoft.com,,,,,,,,,,,,,,,,,,,,,,,
#   remove-moverequest,test4@kc2.extest.microsoft.com,,,,,,,,,0,,,,,,,,,,,,,,

#   Sample2:
#   Operation,Identity,Confirm
#   remove-moverequest,test1@kc2.extest.microsoft.com,0
#   remove-moverequest,test4@kc2.extest.microsoft.com,0
#
#   $recs = gc .\records.csv
#
#   .\BulkMove-Mailbox.ps1 -Records $recs -tenantname 'kc2.extest.microsoft.com' -Test 'false'
#
#
#   Copyright (c) Microsoft Corporation. All rights reserved.
#
#   THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#   OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[string]$BatchName,
	[Parameter(Mandatory = $false)]
	[string]$InputFile,
	[Parameter(Mandatory = $false)]
	[String]$LogFile = $null,
	[String]$IgnoreExcludedFromInitialProvisioning = "False",
	[String]$Test = "True",
	[Parameter(Mandatory=$false)]
	$Records = $null,
	[Parameter(Mandatory=$false)]
	[string] $TenantName = $null
)

if ($(Get-PSSnapin | ?{$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}

function GetObjSearcher
{
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchScope = "SubTree"
    $objSearcher.PropertiesToLoad.Add("msExchOURoot") | out-null
    $objSearcher.PropertiesToLoad.Add("homeMDB") | out-null
    $objSearcher.PropertiesToLoad.Add("msExchArchiveDatabaseLink") | out-null
    $objSearcher.PropertiesToLoad.Add("objectguid") | out-null
    $objSearcher.PropertiesToLoad.Add("msexchmailboxguid") | out-null
    $objSearcher.PropertiesToLoad.Add("msexcharchiveguid") | out-null
    $objSearcher.PropertiesToLoad.Add("msExchOrganizationUpgradeRequest") | out-null
    $RootDN = $(Get-User -resultsize 1).distinguishedname
    $RootDN = $RootDN.substring($RootDN.IndexOf("DC="))
    $objSearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$RootDN")
    return $objSearcher
}

function CreateMoveRequest($ArchiveOnly, $entry, $objectguid, $batchName, $Priority, $CompletedRequestAgeLimit, $PreventCompletion, $PrimaryOnly, $Protect, $BlockFinalization)
{
    $nmrParams = @{}
    $nmrParams.Add("Identity", $objectguid)
    $nmrParams.Add("BatchName", $batchName)
    $nmrParams.Add("Priority", $Priority)
    $nmrParams.Add("CompletedRequestAgeLimit", $CompletedRequestAgeLimit)
    $nmrParams.Add("AllowLargeItems", $true)
    $nmrParams.Add("ErrorAction", "silentlycontinue")
    
    if ($ArchiveOnly)
    {
        $nmrParams.Add("ArchiveOnly", $true)
        if ($entry.TargetDatabase -ne $null -and $entry.TargetDatabase -ne "")
        {   
            $nmrParams.Add("ArchiveTargetDatabase", $entry.TargetDatabase)
        }
    }
    else
    {
        if ($PrimaryOnly)
        {
            $nmrParams.Add("PrimaryOnly", $true)
        }        
        if ($entry.TargetDatabase -ne $null -and $entry.TargetDatabase -ne "")
        {   
            $nmrParams.Add("TargetDatabase", $entry.TargetDatabase)
        }
    }
    
    if($BlockFinalization)
    {
        $nmrParams.Add("SkipMoving", "BlockFinalization")
    }
    
    if($PreventCompletion)
    {
        $nmrParams.Add("PreventCompletion", $true)
    }
    
    if($Protect)
    {
        $nmrParams.Add("Protect", $true)
    }
    
    New-MoveRequest @nmrParams | out-null
}

function GetTargetDBVersion($dbName, $versionMap)
{    
    if($dbName -eq $null -or $dbName -eq "")
    {
        $dbName = "Empty"
        if($versionMap[$dbName] -eq $null)
        {
            if(test-path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup")
            {
                $setupKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup"
            }
            elseif(test-path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup")
            {
                $setupKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup"
            }
            else
            {
                return $null
            }
            $dbversion = "{0}.{1}" -f $($setupKey.MsiProductMajor), $($setupKey.MsiProductMinor)
            $versionMap[$dbName] = $dbversion            
        }
        return $versionMap[$dbName]
    }    
    else
    {
        $dagName = $($dbName.split("-"))[0]    
        if($versionMap[$dagName] -eq $null)
        {
            $dag = Get-DatabaseAvailabilityGroup $dagName
            $server = Get-ExchangeServer $dag.Servers[0].ToString()
            $serverversion = $($server.AdminDisplayVersion.tostring().split(" "))[1]
            $versionMap[$dagName] = $serverversion  
        }
        return $versionMap[$dagName]
    }   
}

function GetOrgInfo($entry)
{
    $orgName = $entry.orgName
    if($orgName -eq $null -or $orgName -eq "")
    {
        return $null
    }
    
    if($script:TenantVersion[$orgName] -eq $null)
    {           
        if($orgname -eq "FirstOrg")
        {
            $script:TenantVersion[$orgName] = "IgNore"
        }
        else
        {
            $org = get-organization $orgName
            if($org -eq $null)
            {
                $script:TenantVersion[$orgName] = "None"
                $Failure = "Tenant_NotFound"
                $failedcount++
                "$(Get-Date),$failedcount,`"$($entry.Guid)`",$objectGuid,`"Failed when checiking prevent completion`", $Failure,$batchName,$($entry.SourceDatabase),$($entry.TargetDatabase),$($entry.ArchiveOnly),$($entry.Overwrite),$($entry.PreventCompletion),$($entry.BlockFinalization)" |  Out-file $failureLog -append              
                return $null
            }
            else
            {
               $versiondetails = $($org.AdminDisplayVersion.tostring().split("("))[1].split(".")  
               $orgversion = "{0}.{1}" -f $versiondetails[0], $versiondetails[1]
               $script:TenantVersion[$orgName] = $orgversion     
               
               if($org.IsUpgradingOrganization) 
               {
                   $script:TenantUpgrade[$orgName] = "Upgrade"
               }
               elseif($org.IsPilotingOrganization -and $org.ServicePlan.tostring() -match "E15.*Pilot")
               {
                   $script:TenantUpgrade[$orgName] = "Pilot"
               }
               else
               {
                   $script:TenantUpgrade[$orgName] = "None"
               }
               return $orgversion
            }
        }               
    }

    return $script:TenantVersion[$orgName]
}

function NeedPreventCompletion($entry)
{   
    if($entry.orgName -eq $null -or $entry.orgName -eq "")
    {
        return $true
    }
    if($script:TenantVersion[$entry.orgName] -eq $null)
    {
        if($(GetOrgInfo $entry) -eq $null)
        {   
            return $true
        }
    }
    if($script:TenantVersion[$entry.orgName] -eq "IgNore")
    {
        return $false
    }
    $targetDbVersion = GetTargetDBVersion $entry.targetdatabase $script:TargetDBs
    $orgVersion = $script:TenantVersion[$entry.orgName]   
    $orgUpgrade = $script:TenantUpgrade[$entry.orgName] 
    #Use Failure column to log the result for version query
    $Failure = "Targetdb version = {0} org version = {1} " -f $targetDbVersion, $orgVersion
    "$(Get-Date),$injected,$count,$($UserList.count),$($entry.Guid),$objectGuid,$Failure,$batchName,$($entry.SourceDatabase),$($entry.TargetDatabase),$ArchiveOnly,$Priority,$($entry.Overwrite),$PreventCompletion,$BlockFinalization" | Out-file $LogFile -append     
    if($targetDbVersion -eq $orgVersion)
    {
        return $false
    }
    elseif($entry.UpgradeRequest -eq "4" -and $orgUpgrade -eq "Pilot")
    {
        return $false
    }
    else
    {
        return $true
    }
}

function AllowOverwriteAll($entry)
{   
    if($entry.orgName -eq $null -or $entry.orgName -eq "")
    {
        return $false
    }
    $orgUpgrade = $script:TenantUpgrade[$entry.orgName] 
    if($script:TenantVersion[$entry.orgName] -eq "None")
    {
        return $false
    }
    elseif($(NeedPreventCompletion $entry) -eq $false -and ($orgUpgrade -eq "Upgrade" -or $entry.orgName -eq "FirstOrg"))
    {   
        return $true                  
    }
    else
    {
        return $false       
    }
}

function RemoveMove($objectguid)
{
    $request = get-moverequest $objectguid -ErrorAction SilentlyContinue
    if($request.requeststyle -eq "IntraOrg")
    {
        Remove-MoveRequest $objectguid -Confirm:$false -ErrorAction SilentlyContinue | out-null
    }
}

function AddStringParam($params, $name, $value)
{
	if($value)
	{
		$params[$name] = $value
	}
}

function AddBoolParam($params, $name, $value)
{
	if($value)
	{
	    if($value -eq '1')
		{
			$params[$name] = $true
		}
		elseif($value -eq '0')
		{
			$params[$name] = $false
		}
	}
}

function IsTenantReadyForCutover($tenantName)
{
	$val = (Get-ExchangeSettings XFR -ConfigName CutOverReadyTenantIds -Force).EffectiveSetting.Value

	if($val -eq $null)
	{
		return $false
	}
	else
	{
		$orgId = (get-organization $tenantName).ExternalDirectoryOrganizationId
		if($val.Contains($orgId ))
		{
			return $true 
		}
		else
		{
			return $false
		}
	}
}

function CheckTenantGLSOverride($tenantName)
{
	LogInfo "Checking the GLSOVerride for the tenant: $($TenantName)"

    $machineName = $env:COMPUTERNAME

	$server = get-ExchangeServer $machineName

	if(!$server -or ($server.serverrole -notmatch 'mailbox')) # if running on a non-mbx machine
	{
		$machineName = (get-exchangeserver | ? {$_.serverrole -match 'mailbox' -and $_.admindisplayversion -match '15.0'})[0].Name
	}
	
	$val = (Get-ExchangeSettings AdDriver -ConfigName GlsTenantOverrides -Server $machineName -Force).EffectiveSetting.Value
	if($val -eq $null)
	{
		LogError "Failed to find the AdDriver.GlsTenantOverrides value scoped to this machine: $machineName"
		exit
	}

	if($val.Contains(":$($TenantName):"))
	{
		LogInfo "Found the tenant in GlsTenantOverrides list of the machine: $machineName"
	}
	else
	{
		LogError "Failed to find the tenant in GlsTenantOverrides list"
		exit
	}	
}

function LogInfo($str)
{
	write-host $str
	$str = "$(get-date -Format 'HH:mm:ss, dd/MM/yyyy') - $($str)"
	$str | Out-file $LogFile -append -force
}

function LogError($str)
{
	$str = "[Error] " + $str
	LogInfo $str
	$str = "$(get-date -Format 'HH:mm:ss, dd/MM/yyyy') - $($str)"
	$str | Out-file $failureLog -append -force
}

function GetParametersString($params)
{
	$paramString = "";
	foreach($key in $params.Keys)
	{
		$paramString += " -" + $key + ":" + $params[$key]
	}

	return $paramString
}

function ContructParams($record)
{
	$params = @{}
	AddStringParam $params 'Identity' $record.Identity
	AddStringParam $params 'RemoteHostName' $record.RemoteHostName
	AddStringParam $params 'RemoteOrganizationName' $record.RemoteOrganizationName
	AddStringParam $params 'TargetDeliveryDomain' $record.TargetDeliveryDomain
	AddStringParam $params 'BadItemLimit' $record.BadItemLimit
	AddStringParam $params 'TargetDatabase' $record.TargetDatabase
	AddStringParam $params 'ArchiveTargetDatabase' $record.ArchiveTargetDatabase
	AddStringParam $params 'BatchName' $record.BatchName
	AddStringParam $params 'LargeItemLimit' $record.LargeItemLimit
	AddStringParam $params 'Priority' $record.Priority
	AddStringParam $params 'StartAfter' $record.StartAfter
	AddStringParam $params 'CompleteAfter' $record.CompleteAfter
	AddStringParam $params 'IncrementalSyncInterval' $record.IncrementalSyncInterval
	AddStringParam $params 'CompletedRequestAgeLimit' $record.CompletedRequestAgeLimit
	AddBoolParam $params 'Protect' $record.Protect
	AddBoolParam $params 'Remote' $record.Remote
	AddBoolParam $params 'PreventCompletion' $record.PreventCompletion
	AddBoolParam $params 'SuspendWhenReadyToComplete' $record.SuspendWhenReadyToComplete
	AddBoolParam $params 'Confirm' $record.Confirm
	AddBoolParam $params 'CheckInitialProvisioningSetting' $record.CheckInitialProvisioningSetting
	AddBoolParam $params 'AllowLargeItems' $record.AllowLargeItems
	AddBoolParam $params 'DoNotPreserveMailboxSignature' $record.DoNotPreserveMailboxSignature
	AddBoolParam $params 'AcceptLargeDataLoss' $record.AcceptLargeDataLoss
	AddBoolParam $params 'ForceOffline' $record.ForceOffline

	if($record.SkipMoving)
	{
		$params['SkipMoving'] = $record.SkipMoving.split(',')
	}

	if($record.InternalFlags)
	{
		$params['InternalFlags'] = $record.InternalFlags.split(',')
	}

	return $params 
}

function InitializeDirectory()
{
    param([string] $path)

    if(-not(Test-Path -Path $path -PathType Container))
    {
        [Void](New-Item -Path $path -ItemType Container)
    }

    $path
}

function GetLogFileName
{
	if (Test-Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup")
	{
		$version = "V14"
	}
	elseif (Test-Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup")
	{
		$version = "V15"
	}
	else
	{
		throw "Can't find version regkey "
	}

	$exchangeInstallPath = (get-item HKLM:\SOFTWARE\Microsoft\ExchangeServer\$version\Setup).GetValue("MsiInstallPath")

	$LogDirectory = InitializeDirectory $(Join-Path $exchangeInstallPath "Logging\XFR")

	$now = [DateTime]::UtcNow.ToString("yyyyMMdd_HHmmss")

	$logFileName = "$LogDirectory\XFR$($now).$([guid]::NewGuid().ToString()).log"

	return $logFileName
}

###############################################################Script Body#######################################################################################################################

if($LogFile -eq $null -or $LogFile -eq '')
{
	$LogFile = GetLogFileName
}

$failureLog = $LogFile + ".failure"

write-Host "Log file: $($LogFile) Failure log file: $($failureLog)"

if($TenantName) # if $TenantName is specified, cross forest move code path gets executed
{
	CheckTenantGLSOverride $tenantName

	$isTenantReadyForCutOver = IsTenantReadyForCutover($TenantName)

	LogInfo "IsTenantReadyForCutOver: $isTenantReadyForCutOver"

	if($Records -eq $null)
	{
		LogError "Null 'Records' argument"
		exit
	}
	else
	{
		$Records = ConvertFrom-Csv $Records 

		if($Records -eq $null -or $Records[0].Operation -eq $null)
		{
			LogError "Invalid 'Records' argument"
			exit
		}

		LogInfo "Processing $($Records.Count) records"

		foreach($record in $Records)
		{	
			$params = ContructParams $record
			$cmdString =  "$($record.Operation) $(GetParametersString($params))"
			if($record.SkipRecord -eq '1')
			{
				LogInfo "Skipping: '$($cmdString)'"
			}
			else
			{
				LogInfo "Executing: '$($cmdString)'"

				$operation = $record.Operation.ToLower();

				if($record.CompleteAfter -ne $null -and ($isTenantReadyForCutOver -eq $false))
				{
					$completeAfter = Get-Date $record.CompleteAfter
					if($completeAfter -lt (Get-Date).AddYears(1))
					{
						LogError "Tenant is not ready for CutOver. New-MoveRequest should be called with CompleteAfter value greater than $((Get-Date).AddYears(1)). Command: '$($cmdString)'"
						continue;
					}
				}
				  
				if(($operation -eq 'new-moverequest') -and ($isTenantReadyForCutOver -eq $false) -and ($record.PreventCompletion -ne '1') -and ($record.SkipMoving -NotMatch 'BlockFinalization'))
				{
					LogError "Tenant is not ready for CutOver. New-MoveRequest should be called with PreventCompletion:true or SkipMoving:BlockFinalization. Command: '$($cmdString)'"
					continue
				}

				if(($operation -eq 'set-moverequest') -and ($record.PreventCompletion -eq '0') -and ($isTenantReadyForCutOver -eq $false))
				{
					LogError "Set-MoveRequest should not be called with PreventCompletion:false if the tenant is not ready for CutOver. Command: '$($cmdString)'"
					continue
				}

				if($Test -ne "True")
				{ 
					$error.clear()
					switch ($operation) 
					{ 
						'new-moverequest' { New-MoveRequest  @params } 
						'set-moverequest' { Set-MoveRequest  @params }
						'remove-moverequest' {Remove-MoveRequest  @params}
						'resume-moverequest' {Resume-MoveRequest  @params}
						default {'Invalid Operation. Name: $($record.Operation)'}
					}

					if ($error.count -gt 0)
					{
						LogError $Error[0]
					}
				}
			}
		}
	}

	LogInfo "Execution completed"
}
else # old BulkMove-Mailbox code is executed
{
	$excludedMDB = @{}
	Get-MailboxDatabase | %{ if ($_.IsExcludedFromProvisioning -or ($_.IsExcludedFromInitialProvisioning -and $IgnoreExcludedFromInitialProvisioning -eq "False")) { $excludedMDB.add($_.Name, 1) }}

	$domain = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
	$ForestMonitor = $domain.split(".")[0] + ".onmicrosoft.com"
	$objSearcher = GetObjSearcher
	$failedcount = 0
	$injected = 0
	$count = 0
	$UserList = @(import-csv $InputFile)

	"Time,Injected,count,Totalcount,InputGuid,Guid,Failure,batchName,SourceDatabase,TargetDatabase,ArchiveOnly,Priority,Overwrite,PreventCompletion,BlockFinalization" | out-file $LogFile
	"Time, TotalFailurecount,MBXGuid,ObjectGuid, Failure, Failuredetail,batchName,SourceDatabase,TargetDatabase,ArchiveOnly,Overwrite,PreventCompletion,BlockFinalization"  | out-file $failureLog 

	$script:TargetDBs = @{}
	$script:TenantVersion = @{}
	$script:TenantUpgrade = @{}

	foreach ($entry in $UserList)
	{
		$count++
		$Failure = ""
		$objectGuid = $entry.Guid
    
		$ArchiveOnly = $false
		if ($entry.ArchiveOnly -eq "TRUE")
		{
			$ArchiveOnly = $true
		}
    
		$PrimaryOnly = $false
		if ($entry.ArchiveOnly -eq "Primary")
		{
			$PrimaryOnly = $true
		}

		$priority = $entry.priority
		if ($priority -eq $null -or $priority -eq "")
		{
			$priority = "Normal"
		}
    
		$CompletedRequestAgeLimit = $entry.CompletedRequestAgeLimit
		if($CompletedRequestAgeLimit -eq $null -or $CompletedRequestAgeLimit -eq "")
		{
			$CompletedRequestAgeLimit = 0
		}

		if ($entry.TargetDatabase -ne $null -AND $excludedMDB.contains($entry.TargetDatabase))
		{
			$Failure = "$($entry.TargetDatabase) cannot be targeted of MBX move"
		}  
		else
		{
			$GuidSearch=""
    		([guid]$objectGuid).tobytearray() | %{ `
        		$temp = [Convert]::ToString($_, 16); `
        		if ($temp.length -eq 1) {$temp = "0" + $temp} `
        		$GuidSearch += "\" + $temp; `
    		}

			$objSearcher.Filter = "(|(objectguid=$GuidSearch)(msExchMailboxGuid=$GuidSearch)(msExchArchiveGUID=$GuidSearch))"

			$colResult = $objSearcher.FindOne()
			if ($colResult -ne $null)
			{
				$objItem = $colResult.GetDirectoryEntry()
    			$objDN = New-Object System.String(@($objItem.distinguishedname))   
				if($objItem.msExchOURoot -eq $null -or $objItem.msExchOURoot.tostring() -eq "")
				{
					if($($objDN.split(","))[1] -eq "CN=Users")
					{
						$entry | add-member -type NoteProperty -Name orgName -Value "FirstOrg"  
					}
				}
				else
				{
					$entry | add-member -type NoteProperty -Name orgName -Value $($objItem.msExchOURoot.ToString().split(","))[0].substring("OU=".length)                               
				}
    			$objectGuid = $(New-Object System.Guid(@($objItem.objectguid))).guid
    			$MBXGuid = New-Object System.Guid(@($objItem.msexchmailboxguid))
    			$ArchiveGuid = New-Object System.Guid(@($objItem.msexcharchiveguid))
				$entry | add-member -type NoteProperty -Name UpgradeRequest -Value $objItem.msExchOrganizationUpgradeRequest
				if ($entry.Guid -eq $MBXGuid.guid)
				{
					$ArchiveOnly = $false
				}
				elseif ($entry.Guid -eq $ArchiveGuid.guid)
				{
					$ArchiveOnly = $true
				}            
            
				# don't load balance monitoring tenants
				if ($objDN -match "exchangemon.net,OU=Microsoft Exchange Hosted Organizations,DC=" -or $objDN -match $ForestMonitor)
				{
					$Failure = "Exclude_ExchangeMon"                
				}            
				elseif ($entry.SourceDatabase -ne $null -and $entry.SourceDatabase -ne "")
				{
					if("$($objItem.homemdb)" -notmatch $entry.SourceDatabase -and "$($objItem.msexcharchivedatabaselink)" -notmatch $entry.SourceDatabase)
					{                
						$Failure = "SourceDatabase_Mismatch"
					}
				}                       
			}
			else
			{
				$Failure = "User_NotFound"
			}
		}
    
		if ($Failure -ne "")
		{
			  $failedcount++
			  "$(Get-Date),$failedcount,`"$($entry.Guid)`",$objectGuid,`"Failed when checiking database matching`", $Failure,$batchName,$($entry.SourceDatabase),$($entry.TargetDatabase),$($entry.ArchiveOnly),$($entry.Overwrite),$($entry.PreventCompletion),$($entry.BlockFinalization)" |  Out-file $failureLog -append              
		}
		elseif ($Test -eq "True")
		{
			sleep -Milliseconds 100
		}
		else
		{
			$RemoveMR = $false
			$PreventCompletion = $false
			$Protect = $false
			$BlockFinalization = $false
			if ($entry.PreventCompletion -eq "True")
			{
				$PreventCompletion = $true
			}
			if ($entry.Delete -eq "True")
			{
				$RemoveMR = $true
			}        
			if ($entry.Protect -eq "True")
			{
				$Protect = $true
			}
			if ($entry.BlockFinalization -eq "True")
			{
				$BlockFinalization = $true
			}
        
			if (!$PreventCompletion -and !$BlockFinalization)
			{            
				$PreventCompletion = NeedPreventCompletion $entry  
				if($PreventCompletion)
				{
					$failedcount++
					$Failure = "org version doesn't match target version and not pilot mode"
					"$(Get-Date),$failedcount,`"$($entry.Guid)`",$objectGuid,`"PreventCompletion or BlockFinalizaiton should be true`", $Failure,$batchName,$($entry.SourceDatabase),$($entry.TargetDatabase),$($entry.ArchiveOnly),$($entry.Overwrite),$PreventCompletion,$BlockFinalization" |  Out-file $failureLog -append
					"$(Get-Date),$injected,$count,$($UserList.count),$($entry.Guid),$objectGuid,$Failure,$batchName,$($entry.SourceDatabase),$($entry.TargetDatabase),$ArchiveOnly,$Priority,$($entry.Overwrite),$PreventCompletion,$BlockFinalization" | Out-file $LogFile -append
					continue
				}              
			}                
        
			$error.clear()
        
			if ($RemoveMR)
			{
				RemoveMove $objectguid
			}
			else
			{
				CreateMoveRequest $ArchiveOnly $entry $objectguid $batchName $Priority $CompletedRequestAgeLimit $PreventCompletion $PrimaryOnly $Protect $BlockFinalization
			}

			if ($error.count -gt 0)
			{
				$Failure = $Error[0].Categoryinfo.Reason
            
				if ($Failure -eq "ManagementObjectAlreadyExistsException")
				{
					$remove = $false               
					$request = get-moverequest $objectguid -ErrorAction SilentlyContinue
					switch ($entry.Overwrite)
					{
						"all"
						{  
							if(AllowOverwriteAll $entry)
							{ 
								$remove = $true
							}
						}
						"allIntra"
						{
							if($request.requeststyle -eq "IntraOrg" -and $(AllowOverwriteAll $entry))
							{
								$remove = $true
							}
						}
						"completedIntra"
						{
							if($request.status -eq "completed" -and $request.requeststyle -eq "IntraOrg")
							{
								$remove = $true
							}
						}
					}
                
					if($remove)
					{              
						Remove-MoveRequest $objectguid -Confirm:$false -ErrorAction SilentlyContinue | out-null                    
					}                                
				}
            
				"$(Get-Date),$injected,$count,$($UserList.count),$($entry.Guid),$objectGuid,$Failure,$batchName,$($entry.SourceDatabase),$($entry.TargetDatabase),$ArchiveOnly,$Priority,$($entry.Overwrite),$PreventCompletion,$BlockFinalization" | Out-file $LogFile -append                

				# repeat the inject 1 more time regardless whether the error has been resolved
				$error.clear()
				if ($RemoveMR)
				{
					RemoveMove $objectguid
				}
				else
				{
					CreateMoveRequest $ArchiveOnly $entry $objectguid $batchName $Priority $CompletedRequestAgeLimit $PreventCompletion $PrimaryOnly $Protect $BlockFinalization
				}

				if ($error.count -gt 0)
				{
					$Failure = $Error[0].Categoryinfo.Reason
					$failedcount++
					"$(Get-Date),$failedcount,`"$($entry.Guid)`",$objectGuid,`"retry still failed`", $Failure,$batchName,$($entry.SourceDatabase),$($entry.TargetDatabase),$($entry.ArchiveOnly),$($entry.Overwrite),$PreventCompletion,$BlockFinalization" |  Out-file $failureLog -append
					$error.clear()
				}
				else
				{
					$Failure = ""
				}
			}
		}

		if ($Failure -eq "")
		{
			$injected++
			$Failure = "Success"
		}

		"$(Get-Date),$injected,$count,$($UserList.count),$($entry.Guid),$objectGuid,$Failure,$batchName,$($entry.SourceDatabase),$($entry.TargetDatabase),$ArchiveOnly,$Priority,$($entry.Overwrite),$PreventCompletion,$BlockFinalization" | Out-file $LogFile -append
	}
}

# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUm4W34e12Y0m7T0RTcmg1dg9d
# MTGgghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI4
# WhcNMTcwNjMwMTkyMTI4WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAipCth86FRu1y
# rtsPu2NLSV7nv6A/oVAcvGrG7VQwRn+wGlXrBf4nyiybHGn9MxuB9u4EMvH8s75d
# kt73WT7lGIT1yCIh9VC9ds1iWfmxHZtYutUOM92+a22ukQW00T8U2yowZ6Gav4Q7
# +9M1UrPniZXDwM3Wqm0wkklmwfgEEm+yyCbMkNRFSCG9PIzZqm6CuBvdji9nMvfu
# TlqxaWbaFgVRaglhz+/eLJT1e45AsGni9XkjKL6VJrabxRAYzEMw4qSWshoHsEh2
# PD1iuKjLvYspWv4EBCQPPIOpGYOxpMWRq0t/gqC+oJnXgHw6D5fZ2Ccqmu4/u3cN
# /aAt+9uw4wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFHbWEvi6BVbwsceywvljICto
# twQRMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBABbNYMMt3JjfMAntjQhHrOz4aUk970f/hJw1jLfYspFpq+Gk
# W3jMkUu3Gev/PjRlr/rDseFIMXEq2tEf/yp72el6cglFB1/cbfDcdimLQD6WPZQy
# AfrpEccCLaouf7mz9DGQ0b9C+ha93XZonTwPqWmp5dc+YiTpeAKc1vao0+ru/fuZ
# ROex8Zd99r6eoZx0tUIxaA5sTWMW6Y+05vZN3Ok8/+hwqMlwgNR/NnVAOg2isk9w
# ox9S1oyY9aRza1jI46fbmC88z944ECfLr9gja3UKRMkB3P246ltsiH1fz0kFAq/l
# 2eurmfoEnhg8n3OHY5a/Zzo0+W9s1ylfUecoZ4UwggYHMIID76ADAgECAgphFmg0
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUXcqk6LH319RCPqEAbZRb6cwUsswwYgYKKwYB
# BAGCNwIBDDFUMFKgKoAoAEIAdQBsAGsATQBvAHYAZQAtAE0AYQBpAGwAYgBvAHgA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAGwjcb0RBKN5ZvJbomB6MRohdkuhpyFTI52Klk4oTCb3
# pj0eZEuTee8JBy1kjBtb0OvEkXDG8lCRxrHEtHgqmHqADtLAv8HixWLdEaJdpA1x
# yCwPefceHt5cVVZntcexL/0XqtDI8slDqUjrWGQ3KRa2LfvlrjrSDf6aGOgTD2DS
# uP+9hpwKm0kI8D1accmjHq9bM1FFPlMQcgYDkc2Wdca8RxttFebHBD2i+RHNt6og
# lhpO65VcMqyplDYo7yxJiNTdDWoA63x6nTM+747w4yZMNMbr2sUrEY5eY4oNAf6a
# Ac1QhojbpoNrDgQIDZSjVnukUT+5aOHlOiG1vBiXHgehggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAA
# marFgZ+Mon2KAAAAAACZMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ0MDBaMCMGCSqGSIb3DQEJ
# BDEWBBT9RrCqzWxGDC0XHFdWZs2V1LFTAzANBgkqhkiG9w0BAQUFAASCAQAKkLRw
# 1R2R6GKZkgAwabEXoT9S478fkRAgpqrcDntuBsQER3BHA/r1E7kiruylOlql2pJ0
# dUQze0Sa9K+aWqCTkkCtgm5l/Oy/vkhiYR/+eRMUdxbA1o+dts0TywWAc4uc5ggd
# rkzAuj7Uisy+un/mFqbX9WISiS/FAn8teB0+9v6/zAFYADoBgAKtOQvUiQHStUeV
# lpzC/MsXFL/B8ieXmoKI2+hilsEplueX13ra8Rdz7JD9x8w+XPYafACCPH1ffQCf
# wH6PiI8qYK1RUhrl6q5sKWOQxpwemYV0VTs4yxGLlq0S6osxfJTl7nUmk5HBMy9+
# tlWsjg/PMauABNLi
# SIG # End signature block
