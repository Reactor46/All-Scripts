## Begin Get-ADDirectReports
Function Get-ADDirectReports{
	<#
	.SYNOPSIS
		This Function retrieve the directreports property from the IdentitySpecified.
		Optionally you can specify the Recurse parameter to find all the indirect
		users reporting to the specify account (Identity).
	
	.DESCRIPTION
		This Function retrieve the directreports property from the IdentitySpecified.
		Optionally you can specify the Recurse parameter to find all the indirect
		users reporting to the specify account (Identity).
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
	
		Blog post: http://www.lazywinadmin.com/2014/10/powershell-who-reports-to-whom-active.html
	
		VERSION HISTORY
		1.0 2014/10/05 Initial Version
	
	.PARAMETER Identity
		Specify the account to inspect
	
	.PARAMETER Recurse
		Specify that you want to retrieve all the indirect users under the account
	
	.EXAMPLE
		Get-ADDirectReports -Identity Test_director
	
Name                SamAccountName      Mail                Manager
----                --------------      ----                -------
test_managerB       test_managerB       test_managerB@la... test_director
test_managerA       test_managerA       test_managerA@la... test_director
		
	.EXAMPLE
		Get-ADDirectReports -Identity Test_director -Recurse
	
Name                SamAccountName      Mail                Manager
----                --------------      ----                -------
test_managerB       test_managerB       test_managerB@la... test_director
test_userB1         test_userB1         test_userB1@lazy... test_managerB
test_userB2         test_userB2         test_userB2@lazy... test_managerB
test_managerA       test_managerA       test_managerA@la... test_director
test_userA2         test_userA2         test_userA2@lazy... test_managerA
test_userA1         test_userA1         test_userA1@lazy... test_managerA
	
	#>
	[CmdletBinding()]
	PARAM (
		[Parameter(Mandatory)]
		[String[]]$Identity,
		[Switch]$Recurse
	)
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction 'Stop' -Verbose:$false }
		}
		CATCH
		{
			Write-Verbose -Message "[BEGIN] Something wrong happened"
			Write-Verbose -Message $Error[0].Exception.Message
		}
	}
	PROCESS
	{
		foreach ($Account in $Identity)
		{
			TRY
			{
				IF ($PSBoundParameters['Recurse'])
				{
					# Get the DirectReports
					Write-Verbose -Message "[PROCESS] Account: $Account (Recursive)"
					Get-Aduser -identity $Account -Properties directreports |
					ForEach-Object -Process {
						$_.directreports | ForEach-Object -Process {
							# Output the current object with the properties Name, SamAccountName, Mail and Manager
							Get-ADUser -Identity $PSItem -Properties * | Select-Object -Property *, @{ Name = "ManagerAccount"; Expression = { (Get-Aduser -identity $psitem.manager).samaccountname } }
							# Gather DirectReports under the current object and so on...
							Get-ADDirectReports -Identity $PSItem -Recurse
						}
					}
				}#IF($PSBoundParameters['Recurse'])
				IF (-not ($PSBoundParameters['Recurse']))
				{
					Write-Verbose -Message "[PROCESS] Account: $Account"
					# Get the DirectReports
					Get-Aduser -identity $Account -Properties directreports | Select-Object -ExpandProperty directReports |
					Get-ADUser -Properties * | Select-Object -Property *, @{ Name = "ManagerAccount"; Expression = { (Get-Aduser -identity $psitem.manager).samaccountname } }
				}#IF (-not($PSBoundParameters['Recurse']))
			}#TRY
			CATCH
			{
				Write-Verbose -Message "[PROCESS] Something wrong happened"
				Write-Verbose -Message $Error[0].Exception.Message
			}
		}
	}
	END
	{
		Remove-Module -Name ActiveDirectory -ErrorAction 'SilentlyContinue' -Verbose:$false | Out-Null
	}
}
## End Get-ADDirectReports
## Begin Get-ADDomains
Function Get-ADDomains{
	$Domains = Get-Domains
	ForEach($Domain in $Domains) 
	{
		$DomainName = $Domain.Name
		$DomainFQDN = ConvertTo-FQDN $DomainName
		
		$ADObject   = [ADSI]"LDAP://$DomainName"
		$sidObject = New-Object System.Security.Principal.SecurityIdentifier( $ADObject.objectSid[ 0 ], 0 )

		Write-Debug "***Get-AdDomains DomName='$DomainName', sidObject='$($sidObject.Value)', name='$DomainFQDN'"

		$Object = New-Object -TypeName PSObject
		$Object | Add-Member -MemberType NoteProperty -Name 'Name'      -Value $DomainFQDN
		$Object | Add-Member -MemberType NoteProperty -Name 'FQDN'      -Value $DomainName
		$Object | Add-Member -MemberType NoteProperty -Name 'ObjectSID' -Value $sidObject.Value
		$Object
	}
}
## End Get-ADDomains
## Begin Get-AdminShare
Function Get-AdminShare {
    [cmdletbinding()]
    Param (
        $Computername = $Computername
    )
    $CIMParams = @{
        Computername = $Computername
        ClassName = 'Win32_Share'
        Property = 'Name', 'Path', 'Description', 'Type'
        ErrorAction = 'Stop'
        Filter = "Type='2147483651' OR Type='2147483646' OR Type='2147483647' OR Type='2147483648'"
    }
    Get-CimInstance @CIMParams | Select-Object Name, Path, Description, 
    @{L='Type';E={$ShareType[[int64]$_.Type]}}
}
## End Get-AdminShare
## Begin Get-ADServices
Function Get-ADServices {
          
    [CmdletBinding()]
    [OutputType([Array])] 
    param
    (
        [Parameter(Position=0, Mandatory = $true, HelpMessage="Provide server names", ValueFromPipeline = $true)]
        $Computername
    )
 
    $ServiceNames = "HealthService","NTDS","NetLogon","DFSR"
    $ErrorActionPreference = "SilentlyContinue"
    $report = @()
 
        $Services = Get-Service -ComputerName $Computername -Name  $ServiceNames
 
        If(!$Services)
        {
            Write-Warning "Something went wrong"
        }
        Else
        {
            # Adding properties to object
            $Object = New-Object PSCustomObject
            $Object | Add-Member -Type NoteProperty -Name "ServerName" -Value $Computername
 
            foreach($item in $Services)
            {
                $Name = $item.Name
                $Object | Add-Member -Type NoteProperty -Name "$Name" -Value $item.Status 
            }
             
            $report += $object
        }
     
    $report
}
## End Get-ADServices
## Begin Get-ADSiteAndSubnet
Function Get-ADSiteAndSubnet {
<#
	.SYNOPSIS
		This Function will retrieve Site names, subnets names and descriptions.

	.DESCRIPTION
		This Function will retrieve Site names, subnets names and descriptions.

	.EXAMPLE
		Get-ADSiteAndSubnet
	
	.EXAMPLE
		Get-ADSiteAndSubnet | Export-Csv -Path .\ADSiteInventory.csv

	.OUTPUTS
		PSObject

	.NOTES
		AUTHOR	: Francois-Xavier Cat
		DATE	: 2014/02/03
		
		HISTORY	:
	
			1.0		2014/02/03	Initial Version
			
	
#>
	[CmdletBinding()]
    PARAM()
    BEGIN {Write-Verbose -Message "[BEGIN] Starting Script..."}
    PROCESS
    {
		TRY{
	        # Domain and Sites Information
	        $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
	        $SiteInfo = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites

	        # Forest Context
	        $ForestType = [System.DirectoryServices.ActiveDirectory.DirectoryContexttype]"forest"
	        $ForestContext = New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext -ArgumentList $ForestType,$Forest
            
            # Distinguished Name of the Configuration Partition
            $Configuration = ([ADSI]"LDAP://RootDSE").configurationNamingContext

            # Get the Subnet Container
            $SubnetsContainer = [ADSI]"LDAP://CN=Subnets,CN=Sites,$Configuration"
            $SubnetsContainerchildren = $SubnetsContainer.Children

	        FOREACH ($item in $SiteInfo){
				
				Write-Verbose -Message "[PROCESS] SITE: $($item.name)"

                $output = @{
                    Name = $item.name
                }
                    FOREACH ($i in $item.Subnets.name){
                        Write-verbose -message "[PROCESS] SUBNET: $i"
                        $output.Subnet = $i
                        $SubnetAdditionalInfo = $SubnetsContainerchildren.Where({$_.name -match $i})

                        Write-verbose -message "[PROCESS] SUBNET: $i - DESCRIPTION: $($SubnetAdditionalInfo.Description)"
                        $output.Description = $($SubnetAdditionalInfo.Description)
                        
                        Write-verbose -message "[PROCESS] OUTPUT INFO"

                        New-Object -TypeName PSObject -Property $output
                    }
	        }#Foreach ($item in $SiteInfo)
		}#TRY
		CATCH
		{
			Write-Warning -Message "[PROCESS] Something Wrong Happened"
			Write-Warning -Message $Error[0]
		}#CATCH
    }#PROCESS
    END
	{
		Write-Verbose -Message "[END] Script Completed!"
	}#END
}
## End Get-ADSiteAndSubnet
## Begin Get-ADSiteInventory
Function Get-ADSiteInventory {
<#
	.SYNOPSIS
		This Function will retrieve information about the Sites and Services of the Active Directory

	.DESCRIPTION
		This Function will retrieve information about the Sites and Services of the Active Directory

	.EXAMPLE
		Get-ADSiteInventory
	
	.EXAMPLE
		Get-ADSiteInventory | Export-Csv -Path .\ADSiteInventory.csv

	.OUTPUTS
		PSObject

	.NOTES
		AUTHOR	: Francois-Xavier Cat
		DATE	: 2014/02/02
		
		HISTORY	:
	
			1.0		2014/02/02	Initial Version
			
	
#>
	[CmdletBinding()]
    PARAM()
    BEGIN {Write-Verbose -Message "[BEGIN] Starting Script..."}
    PROCESS
    {
		TRY{
	        # Domain and Sites Information
	        $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
	        $SiteInfo = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites

	        # Forest Context
	        $ForestType = [System.DirectoryServices.ActiveDirectory.DirectoryContexttype]"forest"
	        $ForestContext = New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext -ArgumentList $ForestType,$Forest
            
            # Distinguished Name of the Configuration Partition
            $Configuration = ([ADSI]"LDAP://RootDSE").configurationNamingContext

            # Get the Subnet Container
            $SubnetsContainer = [ADSI]"LDAP://CN=Subnets,CN=Sites,$Configuration"


	        FOREACH ($item in $SiteInfo){
				
				Write-Verbose -Message "[PROCESS] SITE: $($item.name)"
				
				# Get the Site Links
				Write-Verbose -Message "[PROCESS] SITE: $($item.name) - Getting Site Links"
	            $LinksInfo = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::FindByName($ForestContext,$($item.name))).SiteLinks
				
				# Create PowerShell Object and Output
				Write-Verbose -Message "[PROCESS] SITE: $($item.name) - Preparing Output"

	            New-Object -TypeName PSObject -Property @{
	                Name= $item.Name
                    SiteLinks = $item.SiteLinks -join ","
	                Servers = $item.Servers -join ","
	                Domains = $item.Domains -join ","
	                Options = $item.options
	                AdjacentSites = $item.AdjacentSites -join ','
	                InterSiteTopologyGenerator = $item.InterSiteTopologyGenerator
	                Location = $item.location
                    Subnets = ( $info = Foreach ($i in $item.Subnets.name){
                        $SubnetAdditionalInfo = $SubnetsContainer.Children | Where-Object {$_.name -like "*$i*"}
                        "$i -- $($SubnetAdditionalInfo.Description)" }) -join ","
	                #SiteLinksInfo = $LinksInfo | fl *
	                
	                #SiteLinksInfo = New-Object -TypeName PSObject -Property @{
	                    SiteLinksCost = $LinksInfo.Cost -join ","
	                    ReplicationInterval = $LinksInfo.ReplicationInterval -join ','
	                    ReciprocalReplicationEnabled = $LinksInfo.ReciprocalReplicationEnabled -join ','
	                    NotificationEnabled = $LinksInfo.NotificationEnabled -join ','
	                    TransportType = $LinksInfo.TransportType -join ','
	                    InterSiteReplicationSchedule = $LinksInfo.InterSiteReplicationSchedule -join ','
	                    DataCompressionEnabled = $LinksInfo.DataCompressionEnabled -join ',' 
	                #}
	                #>
	            }#New-Object -TypeName PSoBject
	        }#Foreach ($item in $SiteInfo)
		}#TRY
		CATCH
		{
			Write-Warning -Message "[PROCESS] Something Wrong Happened"
			Write-Warning -Message $Error[0]
		}#CATCH
    }#PROCESS
    END
	{
		Write-Verbose -Message "[END] Script Completed!"
	}#END
}
## End Get-ADSiteInventory
## Begin Get-ADSITokenGroup
Function Get-ADSITokenGroup{
	<#
	.SYNOPSIS
		Retrieve the list of group present in the tokengroups of a user or computer object.
	
	.DESCRIPTION
		Retrieve the list of group present in the tokengroups of a user or computer object.
		TokenGroups attribute
		https://msdn.microsoft.com/en-us/library/ms680275%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
	
	.PARAMETER SamAccountName
		Specifies the SamAccountName to retrieve
	
	.PARAMETER Credential
		Specifies Credential to use
	
	.PARAMETER DomainDistinguishedName
		Specify the Domain or Domain DN path to use
	
	.PARAMETER SizeLimit
		Specify the number of item maximum to retrieve

    .EXAMPLE
        Get-ADSITokenGroup -SamAccountName TestUser

        GroupName            Count SamAccountName
        ---------            ----- --------------
        lazywinadmin\MTL_GroupB     2 TestUser
        lazywinadmin\MTL_GroupA     2 TestUser
        lazywinadmin\MTL_GroupC     2 TestUser
        lazywinadmin\MTL_GroupD     2 TestUser
        lazywinadmin\MTL-GroupE     1 TestUser
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm		
	#>
	[CmdletBinding()]
	param
	(
		[Parameter(ValueFromPipeline = $true)]
		[Alias('UserName', 'Identity')]
		[String]$SamAccountName,
		
		[Alias('RunAs')]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty,
		
		[Alias('DomainDN', 'Domain')]
		[String]$DomainDistinguishedName = $(([adsisearcher]"").Searchroot.path),
		
		[Alias('ResultLimit', 'Limit')]
		[int]$SizeLimit = '100'
	)
	BEGIN
	{
		$GroupList = ""
	}
	PROCESS
	{
		TRY
		{
			# Building the basic search object with some parameters
			$Search = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ErrorAction 'Stop'
			$Search.SizeLimit = $SizeLimit
			$Search.SearchRoot = $DomainDN
			#$Search.Filter = "(&(anr=$SamAccountName))"
			$Search.Filter = "(&((objectclass=user)(samaccountname=$SamAccountName)))"
			
			# Credential
			IF ($PSBoundParameters['Credential'])
			{
				$Cred = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $DomainDistinguishedName, $($Credential.UserName), $($Credential.GetNetworkCredential().password)
				$Search.SearchRoot = $Cred
			}
			
			# Different Domain
			IF ($DomainDistinguishedName)
			{
				IF ($DomainDistinguishedName -notlike "LDAP://*") { $DomainDistinguishedName = "LDAP://$DomainDistinguishedName" }#IF
				Write-Verbose -Message "[PROCESS] Different Domain specified: $DomainDistinguishedName"
				$Search.SearchRoot = $DomainDistinguishedName
			}
			
			$Search.FindAll() | ForEach-Object -Process {
				$Account = $_
				$AccountGetDirectory = $Account.GetDirectoryEntry();
				
				# Add the properties tokenGroups
				$AccountGetDirectory.GetInfoEx(@("tokenGroups"), 0)
				
				
				$($AccountGetDirectory.Get("tokenGroups")) |
				ForEach-Object -Process {
					# Create SecurityIdentifier to translate into group name
					$Principal = New-Object System.Security.Principal.SecurityIdentifier($_, 0)
					
					# Prepare Output
					$Properties = @{
						SamAccountName = $Account.properties.samaccountname -as [string]
						GroupName = $principal.Translate([System.Security.Principal.NTAccount])
					}
					
					# Output Information
					New-Object -TypeName PSObject -Property $Properties
				}
			} | Group-Object -Property groupname |
			ForEach-Object {
				New-Object -TypeName PSObject -Property @{
					SamAccountName = $_.group.samaccountname | Select-Object -Unique
					GroupName = $_.Name
					Count = $_.Count
				}#new-object
			}#Foreach
		}#TRY
		CATCH
		{
			Write-Warning -Message "[PROCESS] Something wrong happened!"
			Write-Warning -Message $error[0].Exception.Message
		}
	}#PROCESS
	END { Write-Verbose -Message "[END] Function Get-ADSITokenGroup End." }
}
## End Get-ADSITokenGroup
## Begin Get-ADSystem
Function Get-ADSystem {
          
    [CmdletBinding()]
    [OutputType([Array])] 
    param
    (
        [Parameter(Position=0, Mandatory = $true, HelpMessage="Provide server names", ValueFromPipeline = $true)]
        $Server
    )
 
    $SystemArray = @()
 
        $Server = $Server.trim()
        $Object = '' | Select ServerName, BootUpTime, UpTime, "Physical RAM", "C: Free Space", "Memory Usage", "CPU usage"
                         
        $Object.ServerName = $Server
 
        # Get OS details using WMI query
        $os = Get-WmiObject win32_operatingsystem -ComputerName $Server -ErrorAction SilentlyContinue | Select-Object LastBootUpTime,LocalDateTime
                         
        If($os)
        {
            # Get bootup time and local date time  
            $LastBootUpTime = [Management.ManagementDateTimeConverter]::ToDateTime(($os).LastBootUpTime)
            $LocalDateTime = [Management.ManagementDateTimeConverter]::ToDateTime(($os).LocalDateTime)
 
            # Calculate uptime - this is automatically a timespan
            $up = $LocalDateTime - $LastBootUpTime
            $uptime = "$($up.Days) days, $($up.Hours)h, $($up.Minutes)mins"
 
            $Object.BootUpTime = $LastBootUpTime
            $Object.UpTime = $uptime
        }
        Else
        {
            $Object.BootUpTime = "(null)"
                $Object.UpTime = "(null)"
        }
 
        # Checking RAM, memory and cpu usage and C: drive free space
        $PhysicalRAM = (Get-WMIObject -class Win32_PhysicalMemory -ComputerName $server | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)})
                         
        If($PhysicalRAM)
        {
            $PhysicalRAM = ("$PhysicalRAM" + " GB")
            $Object."Physical RAM"= $PhysicalRAM
        }
        Else
        {
            $Object.UpTime = "(null)"
        }
    
        $Mem = (Get-WmiObject -Class win32_operatingsystem -ComputerName $Server  | Select-Object @{Name = "MemoryUsage"; Expression = { “{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize)}}).MemoryUsage
                        
        If($Mem)
        {
            $Mem = ("$Mem" + " %")
            $Object."Memory Usage"= $Mem
        }
        Else
        {
            $Object."Memory Usage" = "(null)"
        }
 
        $Cpu =  (Get-WmiObject win32_processor -ComputerName $Server  |  Measure-Object -property LoadPercentage -Average | Select Average).Average 
                         
        If($PhysicalRAM)
        {
            $Cpu = ("$Cpu" + " %")
            $Object."CPU usage"= $Cpu
        }
        Else
        {
            $Object."CPU Usage" = "(null)"
        }
 
        $FreeSpace =  (Get-WmiObject win32_logicaldisk -ComputerName $Server -ErrorAction SilentlyContinue  | Where-Object {$_.deviceID -eq "C:"} | select @{n="FreeSpace";e={[math]::Round($_.FreeSpace/1GB,2)}}).freespace 
                         
        If($FreeSpace)
        {
            $FreeSpace = ("$FreeSpace" + " GB")
            $Object."C: Free Space"= $FreeSpace
        }
        Else
        {
            $Object."C: Free Space" = "(null)"
        }
 
        $SystemArray += $Object
  
        $SystemArray
} 
## End Get-ADSystem
## Begin Get-DCDiag
Function Get-DCDiag {
	<#   
.SYNOPSIS   
   Display DCDiag information on domain controllers.
.DESCRIPTION 
   Display DCDiag information on domain controllers. $adminCredential and $ourDCs should be set externally.
   $ourDCs should be an array of all your domain controllers. This Function will attempt to set it if it is not set via QAD tools.
   $adminCredential should contain a credential object that has access to the DCs. This Function will prompt for credentials if not set.
   If the all dc option is used along side -Type full, it will return an object you can manipulate.
.PARAMETER DC 
    Specify the DC you'd like to run dcdiag on. Use "all" for all DCs.
.PARAMETER Type 
    Specify the type of information you'd like to see. Default is "error". You can specify "full"           
.NOTES   
    Name: Get-DCDiagInfo
    Author: Ginger Ninja (Mike Roberts)
    DateCreated: 12/08/2015
.LINK  
    https://www.gngrninja.com/script-ninja/2015/12/29/powershell-get-dcdiag-commandlet-for-getting-dc-diagnostic-information      
.EXAMPLE   
    Get-DCDiagInfo -DC idcprddc1 -Type full
    $DCDiagInfo = Get-DCDiagInfo -DC all -type full -Verbose
#>  
          
    [CmdletBinding()]
    [OutputType([Array])] 
    param
    (
        [Parameter(Position=0, Mandatory = $true, HelpMessage="Provide server names", ValueFromPipeline = $true)]
        $Computername
    )
    $DCDiagArray = @()
 
            # DCDIAG ===========================================================================================
            $Dcdiag = (Dcdiag.exe /s:$Computername) -split ('[\r\n]')
            $Results = New-Object Object
            $Results | Add-Member -Type NoteProperty -Name "ServerName" -Value $Computername
            $Dcdiag | %{ 
            Switch -RegEx ($_) 
            { 
                "Starting test"      { $TestName   = ($_ -Replace ".*Starting test: ").Trim() } 
                "passed test|failed test" { If ($_ -Match "passed test") {  
                $TestStatus = "Passed" 
                # $TestName 
                # $_ 
                }  
                Else 
                {  
                $TestStatus = "Failed" 
                # $TestName 
                # $_ 
                } 
                } 
            } 
            If ($TestName -ne $Null -And $TestStatus -ne $Null) 
            { 
                $Results | Add-Member -Name $("$TestName".Trim()) -Value $TestStatus -Type NoteProperty -force
                $TestName = $Null; $TestStatus = $Null
            } 
            } 
            $DCDiagArray += $Results
 
    $DCDiagArray
             

    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
        [String]
        $DC,
        
        [Parameter()]
        [ValidateScript({$_ -like "Full" -xor $_ -like "Error"})]
        [String]
        $Type,
        
        [Parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [String]
        $Utility
        )
    
    try {
        
    if (!$ourDCs) {
        
        $ourDCs = Get-ADDomainController -Discover | Select -ExpandProperty Name
    
    }
    
    if (!$adminCredential) {
        
        $adminCredential = Get-Credential -Message "Please enter Domain Admin credentials"
        
    }
    
    Switch ($dc) {
    
    {$_ -eq $null -or $_ -like "*all*" -or $_ -eq ""} {
    
        Switch ($type) {  
            
        {$_ -like "*error*" -or $_ -like $null} {  
             
            [array]$dcErrors = $null
            $i               = 0
            
            foreach ($d in $ourDCs){
            
                $name = $d.Name    
                
                Write-Verbose "Domain controller: $name"
                
                Write-Progress -Activity "Connecting to DC and running dcdiag..." -Status "Current DC: $name" -PercentComplete ($i/$ourDCs.Count*100)
                
                $session = New-PSSession -ComputerName $d.Name -Credential $adminCredential
                
                Write-Verbose "Established PSSession..."
                
                $dcdiag  = Invoke-Command -Session $session -Command  { dcdiag }
                
                Write-Verbose "dcdiag command ran via Invoke-Command..."
            
                if ($dcdiag | ?{$_ -like "*failed test*"}) {
                    
                    Write-Verbose "Failure detected!"
                    $failed = $dcdiag | ?{$_ -like "*failed test*"}
                    Write-Verbose $failed
                    [array]$dcErrors += $failed.Replace(".","").Trim("")
            
                } else {
                
                    $name = $d.Name    
                
                    Write-Verbose "$name passed!"
                    
                }
                
                
                Remove-PSSession -Session $session
                
                Write-Verbose "PSSession closed to: $name"
                $i++
            }
            
            Return $dcErrors
        } 
            
        {$_ -like "*full*"}    {
            
            [array]$dcFull             = $null
            [array]$dcDiagObject       = $null
            $defaultDisplaySet         = 'Name','Error','Diag'
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
            $PSStandardMembers         = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
            $i                         = 0
            
            foreach ($d in $ourDCs){
                
                $diagError = $false
                $name      = $d.Name
                
                Write-Verbose "Domain controller: $name"
                
                Write-Progress -Activity "Connecting to DC and running dcdiag..." -Status "Current DC: $name" -PercentComplete ($i/$ourDCs.Count*100)
                
                $session = New-PSSession -ComputerName $d.Name -Credential $adminCredential
                
                Write-Verbose "Established PSSession..."
                
                $dcdiag  = Invoke-Command -Session $session -Command  { dcdiag }
                
                Write-Verbose "dcdiag command ran via Invoke-Command..."
                
                $diagstring = $dcdiag | Out-String
                
                Write-Verbose $diagstring
                if ($diagstring -like "*failed*") {$diagError = $true}
                
                $dcDiagProperty  = @{Name=$name}
                $dcDiagProperty += @{Error=$diagError}
                $dcDiagProperty += @{Diag=$diagstring}
                $dcO             = New-Object PSObject -Property $dcDiagProperty
                $dcDiagObject   += $dcO
                
                Remove-PSSession -Session $session
                
                Write-Verbose "PSSession closed to: $name"
                
                $i++
            }
            
            $dcDiagObject.PSObject.TypeNames.Insert(0,'User.Information')
            $dcDiagObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
            
            Return $dcDiagObject
        
            }
        
        }
         break         
    }
   
   
    {$_ -notlike "*all*" -or $_ -notlike $null} {
   
        Switch ($type) {
        
        {$_ -like "*error*" -or $_ -like $null} {
        
            if (Get-ADDomainController $dc) { 
    
                Write-Host "Domain controller: " $dc `n -foregroundColor $foregroundColor
            
                $session = New-PSSession -ComputerName $dc -Credential $adminCredential
                $dcdiag  = Invoke-Command -Session $session -Command  { dcdiag }
       
                if ($dcdiag | ?{$_ -like "*failed test*"}) {
                
                    Write-Host "Failure detected!"
                
                    $failed = $dcdiag | ?{$_ -like "*failed test*"}
                
                    Write-Output $failed 
                
                } else { 
                
                    Write-Host $dc " passed!"
                
                }
                    
            Remove-PSSession -Session $session       
            } 
        }
        
        {$_ -like "full"} {
            
            if (Get-ADDomainController $dc) { 
    
                Write-Host "Domain controller: " $dc `n -foregroundColor $foregroundColor
            
                $session = New-PSSession -ComputerName $dc -Credential $adminCredential
                $dcdiag  = Invoke-Command -Session $session -Command  { dcdiag }
                $dcdiag     
                    
                Remove-PSSession -Session $session       
            }     
                
        }
        
    }
    
    }
    
    }
    
    }
    
    Catch  [System.Management.Automation.RuntimeException] {
      
        Write-Warning "Error occured: $_"
    
     }
    
Finally { Write-Verbose "Get-DCDiagInfo Function execution completed."}
#requires -version 2.0
}
## End Get-DCDiag
## Begin Add-ADSubnet
Function Add-ADSubnet{
<#
	.SYNOPSIS
		This Function allow you to add a subnet object in your active directory using ADSI

	.DESCRIPTION
		This Function allow you to add a subnet object in your active directory using ADSI
	
	.PARAMETER  Subnet
		Specifies the Name of the subnet to add

	.PARAMETER  SiteName
		Specifies the Name of the Site where the subnet will be created
	
	.PARAMETER  Description
		Specifies the Description of the subnet

	.PARAMETER  Location
		Specifies the Location of the subnet

	.EXAMPLE
		Add-ADSubnet -Subnet "192.168.10.0/24" -SiteName MTL1
	
	This will create the subnet "192.168.10.0/24" and assign it to the site "MTL1".

	.EXAMPLE
		Add-ADSubnet -Subnet "192.168.10.0/24" -SiteName MTL1 -Description "Workstations VLAN 110" -Location "Montreal, Canada" -verbose
	
	This will create the subnet "192.168.10.0/24" and assign it to the site "MTL1" with the description "Workstations VLAN 110" and the location "Montreal, Canada"
	Using the parameter -Verbose, the script will show the progression of the subnet creation.
	

	.NOTES
		NAME:	FUNCT-AD-SITE-Add-ADSubnet_using_ADSI.ps1
		AUTHOR:	Francois-Xavier CAT 
		DATE:	2013/11/07
		EMAIL:	info@lazywinadmin.com
		WWW:	www.lazywinadmin.com
		TWITTER:@lazywinadm
	
		http://www.lazywinadmin.com/2013/11/powershell-add-ad-site-subnet.html

		VERSION HISTORY:
		1.0 2013.11.07
			Initial Version

#>
	[CmdletBinding()]
	PARAM(
		[Parameter(
			Mandatory=$true,
			Position=1,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="Subnet name to create")]
		[Alias("Name")]
		[String]$Subnet,
		[Parameter(
			Mandatory=$true,
			Position=2,
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="Site to which the subnet will be applied")]
		[Alias("Site")]
		[String]$SiteName,
		[Parameter(
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="Description of the Subnet")]
		[String]$Description,
		[Parameter(
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="Location of the Subnet")]
		[String]$location
	)
	PROCESS{
			TRY{
				$ErrorActionPreference = 'Stop'
				
				# Distinguished Name of the Configuration Partition
				$Configuration = ([ADSI]"LDAP://RootDSE").configurationNamingContext

				# Get the Subnet Container
				$SubnetsContainer = [ADSI]"LDAP://CN=Subnets,CN=Sites,$Configuration"
				
				# Create the Subnet object
				Write-Verbose -Message "$subnet - Creating the subnet object..."
				$SubnetObject = $SubnetsContainer.Create('subnet', "cn=$Subnet")
			
				# Assign the subnet to a site
				$SubnetObject.put("siteObject","cn=$SiteName,CN=Sites,$Configuration")
	
				# Adding the Description information if specified by the user
				IF ($PSBoundParameters['Description']){
					$SubnetObject.Put("description",$Description)
				}
				
				# Adding the Location information if specified by the user
				IF ($PSBoundParameters['Location']){
					$SubnetObject.Put("location",$Location)
				}
				$SubnetObject.setinfo()
				Write-Verbose -Message "$subnet - Subnet added."
			}#TRY
			CATCH{
				Write-Warning -Message "An error happened while creating the subnet: $subnet"
				$error[0].Exception
			}#CATCH
	}#PROCESS Block
	END{
		Write-Verbose -Message "Script Completed"
	}#END Block
}#Function Add-ADSubnet
## End Add-ADSubnet
## Begin Get-ADFSMORole
Function Get-ADFSMORole{
	<#
	.SYNOPSIS
		Retrieve the FSMO Role in the Forest/Domain.
	.DESCRIPTION
		Retrieve the FSMO Role in the Forest/Domain.
	.EXAMPLE
		Get-ADFSMORole
    .EXAMPLE
		Get-ADFSMORole -Credential (Get-Credential -Credential "CONTOSO\SuperAdmin")
    .NOTES
        Francois-Xavier Cat
        www.lazywinadmin.com
        @lazywinadm
		github.com/lazywinadmin
	#>
	[CmdletBinding()]
	PARAM (
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#PARAM
	BEGIN
	{
		TRY
		{
			# Load ActiveDirectory Module if not already loaded.
			IF (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction 'Stop' -Verbose:$false }
		}
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			Write-Warning -Message $Error[0]
		}
	}
	PROCESS
	{
		TRY
		{
            
			IF ($PSBoundParameters['Credential'])
			{
                # Query with the credentials specified
				$ForestRoles = Get-ADForest -Credential $Credential -ErrorAction 'Stop' -ErrorVariable ErrorGetADForest
				$DomainRoles = Get-ADDomain -Credential $Credential -ErrorAction 'Stop' -ErrorVariable ErrorGetADDomain
			}
			ELSE
			{
                # Query with the current credentials
				$ForestRoles = Get-ADForest
				$DomainRoles = Get-ADDomain
			}
			
            # Define Properties
			$Properties = @{
				SchemaMaster = $ForestRoles.SchemaMaster
				DomainNamingMaster = $ForestRoles.DomainNamingMaster
				InfraStructureMaster = $DomainRoles.InfraStructureMaster
				RIDMaster = $DomainRoles.RIDMaster
				PDCEmulator = $DomainRoles.PDCEmulator
			}
			
			New-Object -TypeName PSObject -Property $Properties
		}
		CATCH
		{
			Write-Warning -Message "[PROCESS] Something wrong happened"
			IF ($ErrorGetADForest) { Write-Warning -Message "[PROCESS] Error While retrieving Forest information"}
			IF ($ErrorGetADDomain) { Write-Warning -Message "[PROCESS] Error While retrieving Domain information"}
			Write-Warning -Message $Error[0]
		}
	}#PROCESS
}
## End Get-ADFSMORole

## Active Directory User and Group Functions

## Begin Get-NestedMember
Function Get-NestedMember{
<#
    .SYNOPSIS
        Find all Nested members of a group
    .DESCRIPTION
        Find all Nested members of a group
    .PARAMETER GroupName
        Specify one or more GroupName to audit
    .Example
        Get-NestedMember -GroupName TESTGROUP

        This will find all the indirect members of TESTGROUP
    .Example
        Get-NestedMember -GroupName TESTGROUP,TESTGROUP2

        This will find all the indirect members of TESTGROUP and TESTGROUP2
    .Example
        Get-NestedMember TESTGROUP | Group Name | select name, count

        This will find duplicate

#>
    [CmdletBinding()]
    PARAM(
    [String[]]$GroupName,
    [String]$RelationShipPath,
    [Int]$MaxDepth
    )
    BEGIN 
    {
        $DepthCount = 1

        TRY{
            if(-not(Get-Module Activedirectory -ErrorAction Stop)){
                Write-Verbose -Message "[BEGIN] Loading ActiveDirectory Module"
                Import-Module ActiveDirectory -ErrorAction Stop}
        }
        CATCH
        {
            Write-Warning -Message "[BEGIN] An Error occured"
            Write-Warning -Message $error[0].exception.message
        }
    }
    PROCESS
    {
        TRY
        {
            FOREACH ($Group in $GroupName)
            {
                # Get the Group Information
                $GroupObject = Get-ADGroup -Identity $Group -ErrorAction Stop
 
                IF($GroupObject)
                {
                    # Get the Members of the group
                    $GroupObject | Get-ADGroupMember -ErrorAction Stop | ForEach-Object -Process {
                        
                        # Get the name of the current group (to reuse in output)
                        $ParentGroup = $GroupObject.Name
                        

                        # Avoid circular
                        IF($RelationShipPath -notlike ".\ $($GroupObject.samaccountname) \*")
                        {
                            if($PSBoundParameters["RelationShipPath"]) {
                            
                                $RelationShipPath = "$RelationShipPath \ $($GroupObject.samaccountname)"
                            
                                }
                            Else{$RelationShipPath = ".\ $($GroupObject.samaccountname)"}

                            Write-Verbose -Message "[PROCESS] Name:$($_.name) | ObjectClass:$($_.ObjectClass)"
                            $CurrentObject = $_
                            switch ($_.ObjectClass)
                            {   
                                "group" {
                                    # Output Object
                                    $CurrentObject | Select-Object Name,SamAccountName,ObjectClass,DistinguishedName,@{Label="ParentGroup";Expression={$ParentGroup}}, @{Label="RelationShipPath";Expression={$RelationShipPath}}
                                
                                    if (-not($DepthCount -lt $MaxDepth)){
                                        # Find Child
                                        Get-NestedMember -GroupName $CurrentObject.Name -RelationShipPath $RelationShipPath
                                        $DepthCount++
                                    }
                                }#Group
                                default { $CurrentObject | Select-Object Name,SamAccountName,ObjectClass,DistinguishedName, @{Label="ParentGroup";Expression={$ParentGroup}},@{Label="RelationShipPath";Expression={$RelationShipPath}}}
                            }#Switch
                        }#IF($RelationShipPath -notmatch $($GroupObject.samaccountname))
                        ELSE{Write-Warning -Message "[PROCESS] Circular group membership detected with $($GroupObject.samaccountname)"}
                    }#ForeachObject
                }#IF($GroupObject)
                ELSE {
                    Write-Warning -Message "[PROCESS] Can't find the group $Group"
                }#ELSE
            }#FOREACH ($Group in $GroupName)
        }#TRY
        CATCH{
            Write-Warning -Message "[PROCESS] An Error occured"
            Write-Warning -Message $error[0].exception.message }
    }#PROCESS
    END
    {
        Write-Verbose -Message "[END] Get-NestedMember"
    }
}
## End Get-NestedMember
## Begin Get-AdGroupNestedGroupMembership
Function Get-AdGroupNestedGroupMembership {
##########################################################################################################
<#
.SYNOPSIS
    Get's an Active Directory group's nested group memberships.

.DESCRIPTION
    Displays nested group membership details for all the groups that an AD group is a member of. Searches 
    up through the group membership hiererachy using a slight modification of the script founde here:

    https://blogs.msdn.microsoft.com/b/adpowershell/archive/2009/09/05/token-bloat-troubleshooting-by-analyzing-group-nesting-in-ad.aspx

    Produces a custom object for each group and has an option -TreeView switch for a hierarchical
    display.

.EXAMPLE
    Get-AdGroupNestedGroupMembership -Group "CN=Dreamers,OU=Groups,DC=halo,DC=net" -Domain "halo.net"

    Shows the nested group membership for the group 'CN=Dreamers,OU=Groups,DC=halo,DC=net', from the 
    'halo.net' domain.

    For example:

    BaseGroup                  : CN=Dreamers,OU=Groups,DC=halo,DC=net
    BaseGroupAdminCount        : 1
    MaxNestingLevel            : 2
    NestedGroupMembershipCount : 3
    DistinguishedName          : CN=Super Admins,OU=Groups,DC=corp,DC=halo,DC=net
    GroupCategory              : Security
    GroupScope                 : Universal
    Name                       : Super Admins
    ObjectClass                : group
    ObjectGUID                 : a96c6da9-b45b-4820-b30f-8e7f8a256ca8
    SamAccountName             : Super Admins
    SID                        : S-1-5-21-1741080060-3640901959-2469866113-1106


.EXAMPLE
    Get-AdGroupNestedGroupMembership -Group "Dreamers" -Domain "halo.net" -TreeView 

    Shows the nested group membership for the group 'Dreamers', from the 'halo.net' domain, 
    with a hierarchical tree view.

    For example:

    Super Admins
    +-Enterprise Admins
      +-Denied RODC Password Replication Group
      +-Administrators

    BaseGroup                  : CN=Dreamers,OU=Groups,DC=halo,DC=net
    BaseGroupAdminCount        : 1
    MaxNestingLevel            : 2
    NestedGroupMembershipCount : 3
    DistinguishedName          : CN=Super Admins,OU=Groups,DC=corp,DC=halo,DC=net
    GroupCategory              : Security
    GroupScope                 : Universal
    Name                       : Super Admins
    ObjectClass                : group
    ObjectGUID                 : a96c6da9-b45b-4820-b30f-8e7f8a256ca8
    SamAccountName             : Super Admins
    SID                        : S-1-5-21-1741080060-3640901959-2469866113-1106

.EXAMPLE
    Get-AdGroup -Identity 'Dreamers' | Get-AdGroupNestedGroupMembership -TreeView

    Gets an object for the AD group 'Dreamers' and then pipes it into the Get-AdGroupNestedGroupMembership
    Function. Provides a hierarchical tree view. Uses the current domain.

    For example:

    Super Admins
    +-Enterprise Admins
      +-Denied RODC Password Replication Group
      +-Administrators

    BaseGroup                  : CN=Dreamers,OU=Groups,DC=halo,DC=net
    BaseGroupAdminCount        : 1
    MaxNestingLevel            : 2
    NestedGroupMembershipCount : 3
    DistinguishedName          : CN=Super Admins,OU=Groups,DC=corp,DC=halo,DC=net
    GroupCategory              : Security
    GroupScope                 : Universal
    Name                       : Super Admins
    ObjectClass                : group
    ObjectGUID                 : a96c6da9-b45b-4820-b30f-8e7f8a256ca8
    SamAccountName             : Super Admins
    SID                        : S-1-5-21-1741080060-3640901959-2469866113-1106

.EXAMPLE
    Get-AdGroup -Filter * -SearchBase "OU=Groups,DC=halo,DC=net" -Server 'halo.net' |
    Get-AdGroupNestedGroupMembership | 
    Export-CSV -Path d:\users\timh\nestings.csv

    Gets all of the groups from the 'Groups' OU in the 'halo.net' domain and each AD object into
    the Get-AdGroupNestedGroupMembership Function. Objects from the Get-AdGroupGroupMembership Function
    are then exported to a CSV file named d:\users\timh\nestings.csv

    For example:

    #TYPE Microsoft.ActiveDirectory.Management.AdGroup
    "BaseGroup","BaseGroupAdminCount","MaxNestingLevel","NestedGroupMembershipCount","DistinguishedName"...
    "CN=Dreamers,OU=Groups,DC=halo,DC=net","1","2","3","CN=Super Admins,OU=Groups,DC=corp,DC=halo,DC=net"...

.NOTES
    THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
    FITNESS FOR A PARTICULAR PURPOSE.

    This sample is not supported under any Microsoft standard support program or service. 
    The script is provided AS IS without warranty of any kind. Microsoft further disclaims all
    implied warranties including, without limitation, any implied warranties of merchantability
    or of fitness for a particular purpose. The entire risk arising out of the use or performance
    of the sample and documentation remains with you. In no event shall Microsoft, its authors,
    or anyone else involved in the creation, production, or delivery of the script be liable for 
    any damages whatsoever (including, without limitation, damages for loss of business profits, 
    business interruption, loss of business information, or other pecuniary loss) arising out of 
    the use of or inability to use the sample or documentation, even if Microsoft has been advised 
    of the possibility of such damages, rising out of the use of or inability to use the sample script, 
    even if Microsoft has been advised of the possibility of such damages. 
#>
##########################################################################################################

##################################
## Script Options and Parameters
##################################

#Requires -version 3

#Define and validate parameters
[CmdletBinding()]
Param(
      #The target group
      [parameter(Mandatory,Position=1,ValueFromPipeline=$True)]
      [ValidateScript({Get-AdGroup -Identity $_})] 
      [String]$Group,

      #The target domain
      [parameter(Position=2)]
      [ValidateScript({Get-ADDomain -Identity $_})] 
      [String]$Domain = $(Get-ADDomain).Name,

      #Whether to produce the tree view
      [Switch]$TreeView
      )

#Set strict mode to identify typographical errors (uncomment whilst editing script)
#Set-StrictMode -version Latest

#Version : 1.0


##########################################################################################################


    ########
    ## Main
    ########

    Begin {

        #Load Get-ADNestedGroups Function...

        ##########################################################################################################

        #################################
        ## Function - Get-ADNestedGroups
        #################################

        Function Get-ADNestedGroups {

        <#
        Code adapted from:
        https://blogs.msdn.microsoft.com/b/adpowershell/archive/2009/09/05/token-bloat-troubleshooting-by-analyzing-group-nesting-in-ad.aspx    
        #>

            Param ( 
                [Parameter(Mandatory=$true, 
                    Position=0, 
                    ValueFromPipeline=$true, 
                    HelpMessage="DN or ObjectGUID of the AD Group." 
                )] 
                [string]$groupIdentity, 
                [string]$groupDn,
                [int]$groupAdmin = 0,
                [switch]$showTree 
                ) 

            $global:numberOfRecursiveGroupMemberships = 0 
            $lastGroupAtALevelFlags = @() 

            Function Get-GroupNesting ([string] $identity, [int] $level, [hashtable] $groupsVisitedBeforeThisOne, [bool] $lastGroupOfTheLevel) 
            { 
                $group = $null 
                $group = Get-AdGroup -Identity $identity -Properties "memberOf"
                if($lastGroupAtALevelFlags.Count -le $level) 
                { 
                    $lastGroupAtALevelFlags = $lastGroupAtALevelFlags + 0 
                } 
                if($group -ne $null) 
                { 
                    if($showTree) 
                    { 
                        for($i = 0; $i -lt $level – 1; $i++) 
                        { 
                            if($lastGroupAtALevelFlags[$i] -ne 0) 
                            { 
                                Write-Host -ForegroundColor Yellow -NoNewline "  " 
                            } 
                            else 
                            { 
                                Write-Host -ForegroundColor Yellow -NoNewline "¦ " 
                            } 
                        } 
                        if($level -ne 0) 
                        { 
                            if($lastGroupOfTheLevel) 
                            { 
                                Write-Host -ForegroundColor Yellow -NoNewline "+-" 
                            } 
                            else 
                            { 
                                Write-Host -ForegroundColor Yellow -NoNewline "+-" 
                            } 
                        } 
                        Write-Host -ForegroundColor Yellow $group.Name 
                    } 
                    $groupsVisitedBeforeThisOne.Add($group.distinguishedName,$null) 
                    $global:numberOfRecursiveGroupMemberships ++ 
                    $groupMemberShipCount = $group.memberOf.Count 
                    if ($groupMemberShipCount -gt 0) 
                    { 
                        $maxMemberGroupLevel = 0 
                        $count = 0 
                        foreach($groupDN in $group.memberOf) 
                        { 
                            $count++ 
                            $lastGroupOfThisLevel = $false 
                            if($count -eq $groupMemberShipCount){$lastGroupOfThisLevel = $true; $lastGroupAtALevelFlags[$level] = 1} 
                            if(-not $groupsVisitedBeforeThisOne.Contains($groupDN)) #prevent cyclic dependancies 
                            { 
                                $memberGroupLevel = Get-GroupNesting -Identity $groupDN -Level $($level+1) -GroupsVisitedBeforeThisOne $groupsVisitedBeforeThisOne -lastGroupOfTheLevel $lastGroupOfThisLevel 
                                if ($memberGroupLevel -gt $maxMemberGroupLevel){$maxMemberGroupLevel = $memberGroupLevel} 
                            } 
                        } 
                        $level = $maxMemberGroupLevel 
                    } 
                    else #we’ve reached the top level group, return it’s height 
                    { 
                        return $level 
                    } 
                    return $level 
                } 
            } 
            $global:numberOfRecursiveGroupMemberships = 0 
            $groupObj = $null 
            $groupObj = Get-AdGroup -Identity $groupIdentity
            if($groupObj) 
            { 
                [int]$maxNestingLevel = Get-GroupNesting -Identity $groupIdentity -Level 0 -GroupsVisitedBeforeThisOne @{} -lastGroupOfTheLevel $false 
                Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name BaseGroup -Value $groupDn -Force
                Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name BaseGroupAdminCount -Value $groupAdmin -Force
                Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name MaxNestingLevel -Value $maxNestingLevel -Force 
                Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name NestedGroupMembershipCount -Value $($global:numberOfRecursiveGroupMemberships – 1) -Force 
                $groupObj 
            }

        }   #end of Function Get-ADNestedGroups


        ##########################################################################################################


        #Connect to a Global Catalogue
        $GC = New-PSDrive -PSProvider ActiveDirectory -Server $Domain -Root "" –GlobalCatalog –Name GC

        #Error checking
        if ($GC) {

            #Set location to GC drive
            Set-Location -Path GC:

        }   #end of if ($GC)
        else {

            #Error and exit
            Write-Error -Message "Failed to create GC drive. Exiting Function..."

        }   #end of else ($GC)


    }   #end of Begin


    Process {

        #Now get a list of the group's group memberships
        $AdGroup = Get-AdGroup -Identity $Group -Server $Domain -Properties MemberOf,AdminCount -ErrorAction SilentlyContinue
        $Groups = ($AdGroup).MemberOf

        #Error checking
        if ($Groups) {

            #Loop through each of the groups found
            foreach ($Group in $Groups) {

                #Run group query with or without -TreeView
                if ($TreeView) {

                    #Call Get-ADNestedGroups Function with -showTree
                    Get-ADNestedGroups -groupIdentity $Group -groupDN ($AdGroup).DistinguishedName -groupAdmin ($AdGroup).AdminCount -showTree 


                }   #end of if ($TreeView)
                else {

                    #Call Get-ADNestedGroups Function without -showTree
                    Get-ADNestedGroups -groupIdentity $Group -groupDN ($AdGroup).DistinguishedName -groupAdmin ($AdGroup).AdminCount

                }   #end of else $TreeView


            }   #end of foreach ($Group in $Groups)

        }   #end of if ($Groups)
        else {

            Write-Warning -Message "No group memberships returned for group - $group"

        }   #end of else ($Groups)

    }   #end of Process

    End {

        #Exit the GC PS drive and remove
        if ((Get-Location).Drive.Name -eq "GC") {

            #Move to C: drive
            C:

        }   #end of if ((Get-Location).Drive.Name -eq "GC")

    }   #end of End

}   #end of Function Get-AdGroupNestedGroupMembership
## End Get-AdGroupNestedGroupMembership
## Begin Get-AccountLockedOut
Function Get-AccountLockedOut{
	
<#
.SYNOPSIS
	This Function will find the device where the account get lockedout
.DESCRIPTION
	This Function will find the device where the account get lockedout.
	It will query directly the PDC for this information
	
.PARAMETER DomainName
	Specifies the DomainName to query, by default it takes the current domain ($env:USERDOMAIN)
.PARAMETER UserName
	Specifies the DomainName to query, by default it takes the current domain ($env:USERDOMAIN)
.EXAMPLE
	Get-AccountLockedOut -UserName * -StartTime (Get-Date).AddDays(-5) -Credential (Get-Credential)
	
	This will retrieve the all the users lockedout in the last 5 days using the credential specify by the user.
	It might not retrieve the information very far in the past if the PDC logs are filling up very fast.
	
.EXAMPLE
	Get-AccountLockedOut -UserName "Francois-Xavier.cat" -StartTime (Get-Date).AddDays(-2)
#>
	
	#Requires -Version 3.0
	[CmdletBinding()]
	param (
		[string]$DomainName = $env:USERDOMAIN,
		[Parameter()]
		[ValidateNotNullorEmpty()]
		[string]$UserName = '*',
		[datetime]$StartTime = (Get-Date).AddDays(-1),
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)
	BEGIN
	{
		TRY
		{
            #Variables
            $TimeDifference = (Get-Date) - $StartTime

			Write-Verbose -Message "[BEGIN] Looking for PDC..."
			
			Function Get-PDCServer
			{
	<#
	.SYNOPSIS
		Retrieve the Domain Controller with the PDC Role in the domain
	#>
				PARAM (
					$Domain = $env:USERDOMAIN,
					$Credential = [System.Management.Automation.PSCredential]::Empty
				)
				
				IF ($PSBoundParameters['Credential'])
				{
					
					[System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(
					(New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext -ArgumentList 'Domain', $Domain, $($Credential.UserName), $($Credential.GetNetworkCredential().password))
					).PdcRoleOwner.name
				}#Credentials
				ELSE
				{
					[System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(
					(New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $Domain))
					).PdcRoleOwner.name
				}
			}#Function Get-PDCServer
			
			Write-Verbose -Message "[BEGIN] PDC is $(Get-PDCServer)"
		}#TRY
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			Write-Warning -Message $Error[0]
		}
		
	}#BEGIN
	PROCESS
	{
		TRY
		{
			# Define the parameters
			$Splatting = @{ }
			
			# Add the credential to the splatting if specified
			IF ($PSBoundParameters['Credential'])
			{
                Write-Verbose -Message "[PROCESS] Credential Specified"
				$Splatting.Credential = $Credential
				$Splatting.ComputerName = $(Get-PDCServer -Domain $DomainName -Credential $Credential)
			}
			ELSE
			{
				$Splatting.ComputerName =$(Get-PDCServer -Domain $DomainName)
			}
			
			# Query the PDC
            Write-Verbose -Message "[PROCESS] Querying PDC for LockedOut Account in the last Days:$($TimeDifference.days) Hours: $($TimeDifference.Hours) Minutes: $($TimeDifference.Minutes) Seconds: $($TimeDifference.seconds)"
			Invoke-Command @Splatting -ScriptBlock {
				
				# Query Security Logs
				Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4740; StartTime = $Using:StartTime } |
				Where-Object { $_.Properties[0].Value -like "$Using:UserName" } |
				Select-Object -Property TimeCreated,
							  @{ Label = 'UserName'; Expression = { $_.Properties[0].Value } },
							  @{ Label = 'ClientName'; Expression = { $_.Properties[1].Value } }
			} | Select-Object -Property TimeCreated, UserName, ClientName
		}#TRY
		CATCH
		{
				
		}
	}#PROCESS
}
## End Get-AccountLockedOut
## Begin Get-LoggedOnUser
Function Get-LoggedOnUser ($computername){

$regexa = '.+Domain="(.+)",Name="(.+)"$'
$regexd = '.+LogonId="(\d+)"$'

$logontype = @{
"0"="Local System"
"2"="Interactive" #(Local logon)
"3"="Network" # (Remote logon)
"4"="Batch" # (Scheduled task)
"5"="Service" # (Service account logon)
"7"="Unlock" #(Screen saver)
"8"="NetworkCleartext" # (Cleartext network logon)
"9"="NewCredentials" #(RunAs using alternate credentials)
"10"="RemoteInteractive" #(RDP\TS\RemoteAssistance)
"11"="CachedInteractive" #(Local w\cached credentials)
}

$logon_sessions = @(gwmi win32_logonsession -ComputerName $computername)
$logon_users = @(gwmi win32_loggedonuser -ComputerName $computername)

$session_user = @{}

$logon_users |% {
$_.antecedent -match $regexa > $nul
$username = $matches[1] + "\" + $matches[2]
$_.dependent -match $regexd > $nul
$session = $matches[1]
$session_user[$session] += $username
}

$logon_sessions |%{
$starttime = [management.managementdatetimeconverter]::todatetime($_.starttime)

$loggedonuser = New-Object -TypeName psobject
$loggedonuser | Add-Member -MemberType NoteProperty -Name "Session" -Value $_.logonid
$loggedonuser | Add-Member -MemberType NoteProperty -Name "User" -Value $session_user[$_.logonid]
$loggedonuser | Add-Member -MemberType NoteProperty -Name "Type" -Value $logontype[$_.logontype.tostring()]
$loggedonuser | Add-Member -MemberType NoteProperty -Name "Auth" -Value $_.authenticationpackage
$loggedonuser | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $starttime

$loggedonuser
		}
	}
## End Get-LoggedOnUser
## Begin Get-ADUserLastLogon
Function Get-ADUserLastLogon([string]$userName){
  $dcs = Get-ADDomainController -Filter {Name -like "*"}
  $time = 0
  foreach($dc in $dcs)
  { 
    $hostname = $dc.HostName
    $user = Get-ADUser $userName | Get-ADObject -Properties lastLogon 
    if($user.LastLogon -gt $time) 
    {
      $time = $user.LastLogon
    }
  }
  $dt = [DateTime]::FromFileTime($time)
  Write-Host $username "last logged on at:" $dt }
## End Get-ADUserLastLogon
## Begin Test-ADCredential
Function Test-ADCredential {
	Param($username, $password, $domain)
	Add-Type -AssemblyName System.DirectoryServices.AccountManagement
	$ct = [System.DirectoryServices.AccountManagement.ContextType]::Domain
	$pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($ct, $domain)
	New-Object PSObject -Property @{
		UserName = $username;
		IsValid = $pc.ValidateCredentials($username, $password).ToString()
	}
}
## End Test-ADCredential
## Begin Get-ADUserBadPasswords
Function Get-ADUserBadPasswords {
    [CmdletBinding(
        DefaultParameterSetName = 'All'
    )]
    Param (
        [Parameter(
            ValueFromPipeline = $true,
            ParameterSetName = 'ByUser'
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]$Identity
        ,
        [string]$DomainController = (Get-ADDomain).PDCEmulator
        ,
        [datetime]$StartTime
        ,
        [datetime]$EndTime
    )
    Begin {
        $LogonType = @{
            '2' = 'Interactive'
            '3' = 'Network'
            '4' = 'Batch'
            '5' = 'Service'
            '7' = 'Unlock'
            '8' = 'Networkcleartext'
            '9' = 'NewCredentials'
            '10' = 'RemoteInteractive'
            '11' = 'CachedInteractive'
        }
        $filterHt = @{
            LogName = 'Security'
            ID = 4625
        }
        if ($PSBoundParameters.ContainsKey('StartTime')){
            $filterHt['StartTime'] = $StartTime
        }
        if ($PSBoundParameters.ContainsKey('EndTime')){
            $filterHt['EndTime'] = $EndTime
        }
        # Query the event log just once instead of for each user if using the pipeline
        $events = Get-WinEvent -ComputerName $DomainController -FilterHashtable $filterHt
    }
    Process {
        if ($PSCmdlet.ParameterSetName -eq 'ByUser'){
            $user = Get-ADUser $Identity
            # Filter for the user
            $output = $events | Where-Object {$_.Properties[5].Value -eq $user.SamAccountName}
        } else {
            $output = $events
        }
        foreach ($event in $output){
            [pscustomobject]@{
                TargetAccount = $event.properties.Value[5]
                LogonType = $LogonType["$($event.properties.Value[10])"]
                CallingComputer = $event.Properties.Value[13]
                IPAddress = $event.Properties.Value[19]
                TimeStamp = $event.TimeCreated
            }
        }
    }
    End{}
}
## End Get-ADUserBadPasswords
## Begin Get-PrivilegedGroupsMemberCount
Function Get-PrivilegedGroupsMemberCount{
	Param (
		[Parameter( Mandatory = $true, ValueFromPipeline = $true )]
		$Domains
	)

	## Jeff W. said this was original code, but until I got ahold of it and
	## rewrote it, it looked only slightly changed from:
	## https://gallery.technet.microsoft.com/scriptcenter/List-Membership-In-bff89703
	## So I give them both credit. :-)
	
	## the $Domains param is the output from Get-AdDomains above
	ForEach( $Domain in $Domains ) 
	{
		$DomainSIDValue = $Domain.ObjectSID
		$DomainName     = $Domain.Name
		$DomainFQDN     = $Domain.FQDN

		Write-Debug "***Get-PrivilegedGroupsMemberCount: domainName='$domainName', domainSid='$domainSidValue'"

		## Carefully chosen from a more complete list at:
		## https://support.microsoft.com/en-us/kb/243330
		## Administrator (not a group, just FYI)    - $DomainSidValue-500
		## Domain Admins                            - $DomainSidValue-512
		## Schema Admins                            - $DomainSidValue-518
		## Enterprise Admins                        - $DomainSidValue-519
		## Group Policy Creator Owners              - $DomainSidValue-520
		## BUILTIN\Administrators                   - S-1-5-32-544
		## BUILTIN\Account Operators                - S-1-5-32-548
		## BUILTIN\Server Operators                 - S-1-5-32-549
		## BUILTIN\Print Operators                  - S-1-5-32-550
		## BUILTIN\Backup Operators                 - S-1-5-32-551
		## BUILTIN\Replicators                      - S-1-5-32-552
		## BUILTIN\Network Configuration Operations - S-1-5-32-556
		## BUILTIN\Incoming Forest Trust Builders   - S-1-5-32-557
		## BUILTIN\Event Log Readers                - S-1-5-32-573
		## BUILTIN\Hyper-V Administrators           - S-1-5-32-578
		## BUILTIN\Remote Management Users          - S-1-5-32-580
		
		## FIXME - we report on all these groups for every domain, however
		## some of them are forest wide (thus the membership will be reported
		## in every domain) and some of the groups only exist in the
		## forest root.
		$PrivilegedGroups = "$DomainSidValue-512", "$DomainSidValue-518",
		                    "$DomainSidValue-519", "$DomainSidValue-520",
							"S-1-5-32-544", "S-1-5-32-548", "S-1-5-32-549",
							"S-1-5-32-550", "S-1-5-32-551", "S-1-5-32-552",
							"S-1-5-32-556", "S-1-5-32-557", "S-1-5-32-573",
							"S-1-5-32-578", "S-1-5-32-580"

		ForEach( $PrivilegedGroup in $PrivilegedGroups ) 
		{
			$source = New-Object DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
			$source.SearchScope = 'Subtree'
			$source.PageSize    = 1000
			$source.Filter      = "(objectSID=$PrivilegedGroup)"
			
			Write-Debug "***Get-PrivilegedGroupsMemberCount: LDAP://$DomainName, (objectSid=$PrivilegedGroup)"
			
			$Groups = $source.FindAll()
			ForEach( $Group in $Groups )
			{
				$DistinguishedName = $Group.Properties.Item( 'distinguishedName' )
				$groupName         = $Group.Properties.Item( 'Name' )

				Write-Debug "***Get-PrivilegedGroupsMemberCount: searching group '$groupName'"

				$Source.Filter = "(memberOf:1.2.840.113556.1.4.1941:=$DistinguishedName)"
				$Users = $null
				## CHECK: I don't think a try/catch is necessary here - MBS
				try 
				{
					$Users = $Source.FindAll()
				} 
				catch 
				{
					# nothing
				}
				If( $null -eq $users )
				{
					## Obsolete: F-I-X-M-E: we should probably Return a PSObject with a count of zero
					## Write-ToCSV and Write-ToWord understand empty Return results.

					Write-Debug "***Get-PrivilegedGroupsMemberCount: no members found in $groupName"
				}
				Else 
				{
					Function GetProperValue
					{
						Param(
							[Object] $object
						)

						If( $object -is [System.DirectoryServices.SearchResultCollection] )
						{
							Return $object.Count
						}
						If( $object -is [System.DirectoryServices.SearchResult] )
						{
							Return 1
						}
						If( $object -is [Array] )
						{
							Return $object.Count
						}
						If( $null -eq $object )
						{
							Return 0
						}

						Return 1
					}

 					[int]$script:MemberCount = GetProperValue $Users

					Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' user count before first filter $MemberCount"

					$Object = New-Object -TypeName PSObject
					$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
					$Object | Add-Member -MemberType NoteProperty -Name 'Group'  -Value $groupName

					$Members = $Users | Where-Object { $_.Properties.Item( 'objectCategory' ).Item( 0 ) -like 'cn=person*' }
					$script:MemberCount = GetProperValue $Members

					Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' user count after first filter $MemberCount"

					Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' has $MemberCount members"

					$Object | Add-Member -MemberType NoteProperty -Name 'Members' -Value $MemberCount
					$Object
				}
			}
		}
	}
}
## End Get-PrivilegedGroupsMemberCount
## Begin Get-UserLogin
Function Get-UserLogon {
 
    [CmdletBinding()]
 
    param
 
    (
 
        [Parameter ()]
        [String]$Computer,
 
        [Parameter ()]
        [String]$OU,
 
        [Parameter ()]
        [Switch]$All
 
    )
 
    $ErrorActionPreference = "SilentlyContinue"
 
    $result = @()
 
    If ($Computer) {
 
        Invoke-Command -ComputerName $Computer -ScriptBlock { quser } | Select-Object -Skip 1 | Foreach-Object {
 
            $b = $_.trim() -replace '\s+', ' ' -replace '>', '' -split '\s'
 
            If ($b[2] -like 'Disc*') {
 
                $array = ([ordered]@{
                        'User'     = $b[0]
                        'Computer' = $Computer
                        'Date'     = $b[4]
                        'Time'     = $b[5..6] -join ' '
                    })

 
                $result += New-Object -TypeName PSCustomObject -Property $array
 
            }

 
            else {
 
                $array = ([ordered]@{
                        'User'     = $b[0]
                        'Computer' = $Computer
                        'Date'     = $b[5]
                        'Time'     = $b[6..7] -join ' '
                    })

 
                $result += New-Object -TypeName PSCustomObject -Property $array
 
            }

        }

    }

 
    If ($OU) {
 
        $comp = Get-ADComputer -Filter * -SearchBase "$OU" -Properties operatingsystem
 
        $count = $comp.count
 
        If ($count -gt 20) {
 
            Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer"
 
        }

 
        foreach ($u in $comp) {
 
            Invoke-Command -ComputerName $u.Name -ScriptBlock { quser } | Select-Object -Skip 1 | ForEach-Object {
 
                $a = $_.trim() -replace '\s+', ' ' -replace '>', '' -split '\s'
 
                If ($a[2] -like '*Disc*') {
 
                    $array = ([ordered]@{
                            'User'     = $a[0]
                            'Computer' = $u.Name
                            'Date'     = $a[4]
                            'Time'     = $a[5..6] -join ' '
                        })

 
                    $result += New-Object -TypeName PSCustomObject -Property $array
                }

 
                else {
 
                    $array = ([ordered]@{
                            'User'     = $a[0]
                            'Computer' = $u.Name
                            'Date'     = $a[5]
                            'Time'     = $a[6..7] -join ' '
                        })

 
                    $result += New-Object -TypeName PSCustomObject -Property $array
                }

 
            }

 
        }

 
    }

 
    If ($All) {
 
        $comp = Get-ADComputer -Filter * -Properties operatingsystem
 
        $count = $comp.count
 
        If ($count -gt 20) {
 
            Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer ..."
 
        }

 
        foreach ($u in $comp) {
 
            Invoke-Command -ComputerName $u.Name -ScriptBlock { quser } | Select-Object -Skip 1 | ForEach-Object {
 
                $a = $_.trim() -replace '\s+', ' ' -replace '>', '' -split '\s'
 
                If ($a[2] -like '*Disc*') {
 
                    $array = ([ordered]@{
                            'User'     = $a[0]
                            'Computer' = $u.Name
                            'Date'     = $a[4]
                            'Time'     = $a[5..6] -join ' '
                        })

 
                    $result += New-Object -TypeName PSCustomObject -Property $array
 
                }

                else {
 
                    $array = ([ordered]@{
                            'User'     = $a[0]
                            'Computer' = $u.Name
                            'Date'     = $a[5]
                            'Time'     = $a[6..7] -join ' '
                        })

                    $result += New-Object -TypeName PSCustomObject -Property $array
 
                }

 
            }

 
        }

    }

    Write-Output $result
}
## End Get-UserLogin