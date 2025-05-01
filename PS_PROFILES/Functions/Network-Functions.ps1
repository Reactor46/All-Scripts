## Begin Get-WifiSSID
Function Get-WifiSSID{

    <#
    .SYNOPSIS
        Retrieve Wifi SSID and Connection mode information.

    .DESCRIPTION
        Get-WifiSSID Function will check Network Apapter name to get GUID for XML configuration of Wifi.
        You can use Wildcard for adaptor name and cmdlet will get all SSID name belongs to wi-fi name passed.

    .PARAMETER WifiAdaptorName
        String name to specify Wifi Adaptor Name.
        You can use Wildcard to obtain a number of adaptors.
        Not allowed to use regex but can use * for wildcard.

        If you not specified any adaptor name, then defaul name will be use.

    .INPUTS
        system.string

    .OUTPUTS
        system.object
        
    .NOTES
        Author: guitarrapc
        Date:   June 17, 2013

    .EXAMPLE
        C:\PS> Get-WifiSSID

        FileName       : C:\ProgramData\Microsoft\Wlansvc\Profiles\Interfaces\{D43ADEDC-E07D-4B72-98EF-xxxxxx}\{xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx}.xml
        WifiName       : wifiname
        ConnectionMode : auto
        SSIDName       : wifiname
        SSIDHex        : FFFFFFFFFFFFFFFFFFFFFF

    .EXAMPLE
        C:\PS> Get-WifiSSID -WifiAdaptorName "Wi-fi Sample"

        FileName       : C:\ProgramData\Microsoft\Wlansvc\Profiles\Interfaces\{D43ADEDC-E07D-4B72-98EF-xxxxxx}\{xxxxxxx-xxxx-xxxx-xxxx-yyyyyyyyyyy}.xml
        WifiName       : wifiname2
        ConnectionMode : auto
        SSIDName       : wifiname2
        SSIDHex        : FFFFFFFFFFFFFFFFFFFFFC
        
    #>

    [CmdletBinding()]
    param(
        [parameter(
            position = 0,
            mandatory = 0,
            ValueFromPipeLine,
            ValueFromPipeLinebyPropertyName,
            HelpMessage="Specify a Wi-fi Adaptor Name in Network Adaptor list. default : wi-fi*"
        )]
        [string]
        $WifiAdaptorName = "wi-fi*"
    )

    begin
    {
    }

    process
    {

        Write-Verbose "obrain wi-fi GUID where's name contain '$WifiAdaptorName' : Default value is 'wi-fi*' "
        $WifiGUIDs = (Get-NetAdapter -Name $WifiAdaptorName).InterfaceGuid
        
        Write-Verbose "Only run command when GUID was found with AdapterName '$WifiAdaptorName'."
        if (-not($null -eq $WifiGUIDs))
        {
            $InsterfacePath = "C:\ProgramData\Microsoft\Wlansvc\Profiles\Interfaces\"
            foreach ($WifiGUID in $WifiGUIDs)
            {
                $WifiPath = Join-Path $InsterfacePath $WifiGUID

                Write-Verbose "Checking WifiPath is existing or not at '$WifiPath'"
                if (Test-Path $WifiPath)
                {
                    $WifiXmls = Get-ChildItem -Path $WifiPath -Recurse

                    foreach ($wifixml in $WifiXmls)
                    {
                        [xml]$x = Get-Content -Path $wifixml.FullName

                        [PSCustomObject]@{
                        FileName = $WifiXml.FullName
                        WifiName = $x.WLANProfile.Name
                        ConnectionMode = $x.WLANProfile.ConnectionMode
                        SSIDName = $x.WLANProfile.SSIDConfig.SSID.Name
                        SSIDHex = $x.WLANProfile.SSIDConfig.SSID.Hex
                        }
                    }
                }
                else
                {
                    Write-Verbose "Network adaptor was found, but xml was not exist."
                    throw "Wifi GUID Folder not found in $WifiPath!!"
                }
            }
        }
    }

    end
    {
    }

}
## End Get-WifiSSID
## Begin Get-NetStat
Function Get-NetStat{
<#
.SYNOPSIS
	This Function will get the output of netstat -n and parse the output
.DESCRIPTION
	This Function will get the output of netstat -n and parse the output
.LINK
	http://www.lazywinadmin.com/2014/08/powershell-parse-this-netstatexe.html
.NOTES
	Francois-Xavier Cat
	www.lazywinadmin.com
	@LazyWinAdm
#>
	PROCESS
	{
		# Get the output of netstat
		$data = netstat -n
		
		# Keep only the line with the data (we remove the first lines)
		$data = $data[4..$data.count]
		
		# Each line need to be splitted and get rid of unnecessary spaces
		foreach ($line in $data)
		{
			# Get rid of the first whitespaces, at the beginning of the line
			$line = $line -replace '^\s+', ''
			
			# Split each property on whitespaces block
			$line = $line -split '\s+'
			
			# Define the properties
			$properties = @{
				Protocole = $line[0]
				LocalAddressIP = ($line[1] -split ":")[0]
				LocalAddressPort = ($line[1] -split ":")[1]
				ForeignAddressIP = ($line[2] -split ":")[0]
				ForeignAddressPort = ($line[2] -split ":")[1]
				State = $line[3]
			}
			
			# Output the current line
			New-Object -TypeName PSObject -Property $properties
		}
	}
}
## End Get-NetStat
## Begin Get-NetworkInfo
Function Get-NetworkInfo {
    <#   
        .SYNOPSIS   
            Retrieves the network configuration from a local or remote client.      
             
        .DESCRIPTION   
            Retrieves the network configuration from a local or remote client.        
        
        .PARAMETER Computername
            A single or collection of systems to perform the query against
        
        .PARAMETER Credential
            Alternate credentials to use for query of network information        
        
        .PARAMETER Throttle
            Number of asynchonous jobs that will run at a time
        
        .NOTES   
            Name: Get-NetworkInfo.ps1
            Author: Boe Prox
            Version: 1.0
        
        .EXAMPLE 
             Get-NetworkInfo -Computername 'System1'
            
            NICDescription : Ethernet Network Adapter
            MACAddress     : 00:11:22:33:aa:bb
            NICName        : enthad
            Computername   : System1.domain.com
            DHCPEnabled    : True
            WINSPrimary    : 192.0.0.25
            SubnetMask     : {255.255.255.255}
            WINSSecondary  : 192.0.0.26
            DNSServer      : {192.0.0.31, 192.0.0.30}
            IPAddress      : {192.0.0.5}
            DefaultGateway : {192.0.0.1}         
             
            Description 
            ----------- 
            Retrieves the network information from 'System1'      

        .EXAMPLE
            $Servers = Get-Content Servers.txt
            $Servers | Get-NetworkInfo -Throttle 10
            
            Description
            -----------
            Retrieves all of network information from the remote servers while running 10 runspace jobs at a time.  
            
        .EXAMPLE
            (Get-Content Servers.txt) | Get-NetworkInfo -Credential domain\adminuser -Throttle 10
            
            Description
            -----------
            Gathers all of the network information from the systems in the text file. Also uses alternate administrator credentials provided.                                            
    #>
    #Requires -Version 2.0
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('CN','__Server','IPAddress','Server')]
        [string[]]$Computername = $Env:Computername,
        
        [parameter()]
        [Alias('RunAs')]
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty,       
        
        [parameter()]
        [int]$Throttle = 15
    )
    Begin {
        #Function that will be used to process runspace jobs
        Function Get-RunspaceData {
            [cmdletbinding()]
            param(
                [switch]$Wait
            )
            Do {
                $more = $false         
                Foreach($runspace in $runspaces) {
                    If ($runspace.Runspace.isCompleted) {
                        $runspace.powershell.EndInvoke($runspace.Runspace)
                        $runspace.powershell.dispose()
                        $runspace.Runspace = $null
                        $runspace.powershell = $null
                        $Script:i++                  
                    } ElseIf ($runspace.Runspace -ne $null) {
                        $more = $true
                    }
                }
                If ($more -AND $PSBoundParameters['Wait']) {
                    Start-Sleep -Milliseconds 100
                }   
                #Clean out unused runspace jobs
                $temphash = $runspaces.clone()
                $temphash | Where {
                    $_.runspace -eq $Null
                } | ForEach {
                    Write-Verbose ("Removing {0}" -f $_.computer)
                    $Runspaces.remove($_)
                }             
            } while ($more -AND $PSBoundParameters['Wait'])
        }
            
        Write-Verbose ("Performing inital Administrator check")
        $usercontext = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        $IsAdmin = $usercontext.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")                   
        
        #Main collection to hold all data returned from runspace jobs
        $Script:report = @()    
        
        Write-Verbose ("Building hash table for WMI parameters")
        $WMIhash = @{
            Class = "Win32_NetworkAdapterConfiguration"
            Filter = "IPEnabled='$True'"
            ErrorAction = "Stop"
        } 
        
        #Supplied Alternate Credentials?
        If ($PSBoundParameters['Credential']) {
            $wmihash.credential = $Credential
        }
        
        #Define hash table for Get-RunspaceData Function
        $runspacehash = @{}

        #Define Scriptblock for runspaces
        $scriptblock = {
            Param (
                $Computer,
                $wmihash
            )           
            Write-Verbose ("{0}: Checking network connection" -f $Computer)
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                #Check if running against local system and perform necessary actions
                Write-Verbose ("Checking for local system")
                If ($Computer -eq $Env:Computername) {
                    $wmihash.remove('Credential')
                } Else {
                    $wmihash.Computername = $Computer
                }
                Try {
                        Get-WmiObject @WMIhash | ForEach {
                            $IpHash =  @{
                                Computername = $_.DNSHostName
                                DNSDomain = $_.DNSDomain
                                IPAddress = $_.IpAddress
                                SubnetMask = $_.IPSubnet
                                DefaultGateway = $_.DefaultIPGateway
                                DNSServer = $_.DNSServerSearchOrder
                                DHCPEnabled = $_.DHCPEnabled
                                MACAddress  = $_.MACAddress
                                WINSPrimary = $_.WINSPrimaryServer
                                WINSSecondary = $_.WINSSecondaryServer
                                NICName = $_.ServiceName
                                NICDescription = $_.Description
                            }
                            $IpStack = New-Object PSObject -Property $IpHash
                            #Add a unique object typename
                            $IpStack.PSTypeNames.Insert(0,"IPStack.Information")
                            $IpStack 
                        }
                    } Catch {
                        Write-Warning ("{0}: {1}" -f $Computer,$_.Exception.Message)
                        Break
                }
            } Else {
                Write-Warning ("{0}: Unavailable!" -f $Computer)
                Break
            }        
        }
        
        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
        $runspacepool.Open()  
        
        Write-Verbose ("Creating empty collection to hold runspace jobs")
        $Script:runspaces = New-Object System.Collections.ArrayList        
    }
    Process {        
        $totalcount = $computername.count
        Write-Verbose ("Validating that current user is Administrator or supplied alternate credentials")        
        If (-Not ($Computername.count -eq 1 -AND $Computername[0] -eq $Env:Computername)) {
            #Now check that user is either an Administrator or supplied Alternate Credentials
            If (-Not ($IsAdmin -OR $PSBoundParameters['Credential'])) {
                Write-Warning ("You must be an Administrator to perform this action against remote systems!")
                Break
            }
        }
        ForEach ($Computer in $Computername) {
           #Create the powershell instance and supply the scriptblock with the other parameters 
           $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($wmihash)
           
           #Add the runspace into the powershell instance
           $powershell.RunspacePool = $runspacepool
           
           #Create a temporary collection for each runspace
           $temp = "" | Select-Object PowerShell,Runspace,Computer
           $Temp.Computer = $Computer
           $temp.PowerShell = $powershell
           
           #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
           $temp.Runspace = $powershell.BeginInvoke()
           Write-Verbose ("Adding {0} collection" -f $temp.Computer)
           $runspaces.Add($temp) | Out-Null
           
           Write-Verbose ("Checking status of runspace jobs")
           Get-RunspaceData @runspacehash
        }                        
    }
    End {                     
        Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f (@(($runspaces | Where {$_.Runspace -ne $Null}).Count)))
        $runspacehash.Wait = $true
        Get-RunspaceData @runspacehash
        
        Write-Verbose ("Closing the runspace pool")
        $runspacepool.close()               
    }
}
## End Get-NetworkInfo
## Begin Get-WTFismyIP
Function Get-WTFismyIP {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory = $false, HelpMessage = "Return the result as an object")]
        [switch] $AsObject,

        [Parameter(Mandatory = $false, HelpMessage = "Be polite")]
        [switch] $Polite,

        [Parameter(Mandatory = $false, HelpMessage = "Timeout in seconds")]
        [int] $TimeoutSeconds = 5
    )
    
    begin { }
    
    process {
        try {
            $WTFismyIP = Invoke-RestMethod -Method Get -Uri "https://wtfismyip.com/json" -TimeoutSec $TimeoutSeconds

            if ($AsObject.IsPresent) {
                return $WTFismyIP
            }

            $fucking = $polite.IsPresent ? "" : " fucking"

            $properties = [ordered]@{
                "Your$($fucking) IP address"   = $WTFismyIP.YourFuckingIPAddress
                "Your$($fucking) location"     = $WTFismyIP.YourFuckingLocation
                "Your$($fucking) host name"    = $WTFismyIP.YourFuckingHostname
                "Your$($fucking) ISP"          = $WTFismyIP.YourFuckingISP
                "Your$($fucking) tor exit"     = $WTFismyIP.YourFuckingTorExit
                "Your$($fucking) country code" = $WTFismyIP.YourFuckingCountryCode
            }
            
            $obj = New-Object -TypeName psobject -Property $properties

            Write-Output -InputObject $obj
        }
        catch {
            Write-Error -Message "$_"
        }
    }

    end { }
}
## End Get-WTFismyIP
## Begin Invoke-Ping
Function Invoke-Ping{
<#
.SYNOPSIS
    Ping or test connectivity to systems in parallel
    
.DESCRIPTION
    Ping or test connectivity to systems in parallel

    Default action will run a ping against systems
        If Quiet parameter is specified, we return an array of systems that responded
        If Detail parameter is specified, we test WSMan, RemoteReg, RPC, RDP and/or SMB

.PARAMETER ComputerName
    One or more computers to test

.PARAMETER Quiet
    If specified, only return addresses that responded to Test-Connection

.PARAMETER Detail
    Include one or more additional tests as specified:
        WSMan      via Test-WSMan
        RemoteReg  via Microsoft.Win32.RegistryKey
        RPC        via WMI
        RDP        via port 3389
        SMB        via \\ComputerName\C$
        *          All tests

.PARAMETER Timeout
    Time in seconds before we attempt to dispose an individual query.  Default is 20

.PARAMETER Throttle
    Throttle query to this many parallel runspaces.  Default is 100.

.PARAMETER NoCloseOnTimeout
    Do not dispose of timed out tasks or attempt to close the runspace if threads have timed out

    This will prevent the script from hanging in certain situations where threads become non-responsive, at the expense of leaking memory within the PowerShell host.

.EXAMPLE
    Invoke-Ping Server1, Server2, Server3 -Detail *

    # Check for WSMan, Remote Registry, Remote RPC, RDP, and SMB (via C$) connectivity against 3 machines

.EXAMPLE
    $Computers | Invoke-Ping

    # Ping computers in $Computers in parallel

.EXAMPLE
    $Responding = $Computers | Invoke-Ping -Quiet
    
    # Create a list of computers that successfully responded to Test-Connection

.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Invoke-Ping-Test-in-b553242a

.FunctionALITY
    Computers
	
.NOTES
	Warren F
	http://ramblingcookiemonster.github.io/Invoke-Ping/

#>
	[cmdletbinding(DefaultParameterSetName = 'Ping')]
	param (
		[Parameter(ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[string[]]$ComputerName,
		
		[Parameter(ParameterSetName = 'Detail')]
		[validateset("*", "WSMan", "RemoteReg", "RPC", "RDP", "SMB")]
		[string[]]$Detail,
		
		[Parameter(ParameterSetName = 'Ping')]
		[switch]$Quiet,
		
		[int]$Timeout = 20,
		
		[int]$Throttle = 100,
		
		[switch]$NoCloseOnTimeout
	)
	Begin
	{
		
		#http://gallery.technet.microsoft.com/Run-Parallel-Parallel-377fd430
		Function Invoke-Parallel
		{
			[cmdletbinding(DefaultParameterSetName = 'ScriptBlock')]
			Param (
				[Parameter(Mandatory = $false, position = 0, ParameterSetName = 'ScriptBlock')]
				[System.Management.Automation.ScriptBlock]$ScriptBlock,
				
				[Parameter(Mandatory = $false, ParameterSetName = 'ScriptFile')]
				[ValidateScript({ test-path $_ -pathtype leaf })]
				$ScriptFile,
				
				[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
				[Alias('CN', '__Server', 'IPAddress', 'Server', 'ComputerName')]
				[PSObject]$InputObject,
				
				[PSObject]$Parameter,
				
				[switch]$ImportVariables,
				
				[switch]$ImportModules,
				
				[int]$Throttle = 20,
				
				[int]$SleepTimer = 200,
				
				[int]$RunspaceTimeout = 0,
				
				[switch]$NoCloseOnTimeout = $false,
				
				[int]$MaxQueue,
				
				[validatescript({ Test-Path (Split-Path $_ -parent) })]
				[string]$LogFile = "C:\temp\log.log",
				
				[switch]$Quiet = $false
			)
			
			Begin
			{
				
				#No max queue specified?  Estimate one.
				#We use the script scope to resolve an odd PowerShell 2 issue where MaxQueue isn't seen later in the Function
				if (-not $PSBoundParameters.ContainsKey('MaxQueue'))
				{
					if ($RunspaceTimeout -ne 0) { $script:MaxQueue = $Throttle }
					else { $script:MaxQueue = $Throttle * 3 }
				}
				else
				{
					$script:MaxQueue = $MaxQueue
				}
				
				Write-Verbose "Throttle: '$throttle' SleepTimer '$sleepTimer' runSpaceTimeout '$runspaceTimeout' maxQueue '$maxQueue' logFile '$logFile'"
				
				#If they want to import variables or modules, create a clean runspace, get loaded items, use those to exclude items
				if ($ImportVariables -or $ImportModules)
				{
					$StandardUserEnv = [powershell]::Create().addscript({
						
						#Get modules and snapins in this clean runspace
						$Modules = Get-Module | Select -ExpandProperty Name
						$Snapins = Get-PSSnapin | Select -ExpandProperty Name
						
						#Get variables in this clean runspace
						#Called last to get vars like $? into session
						$Variables = Get-Variable | Select -ExpandProperty Name
						
						#Return a hashtable where we can access each.
						@{
							Variables = $Variables
							Modules = $Modules
							Snapins = $Snapins
						}
					}).invoke()[0]
					
					if ($ImportVariables)
					{
						#Exclude common parameters, bound parameters, and automatic variables
						Function _temp { [cmdletbinding()]
							param () }
						$VariablesToExclude = @((Get-Command _temp | Select -ExpandProperty parameters).Keys + $PSBoundParameters.Keys + $StandardUserEnv.Variables)
						Write-Verbose "Excluding variables $(($VariablesToExclude | sort) -join ", ")"
						
						# we don't use 'Get-Variable -Exclude', because it uses regexps. 
						# One of the veriables that we pass is '$?'. 
						# There could be other variables with such problems.
						# Scope 2 required if we move to a real module
						$UserVariables = @(Get-Variable | Where { -not ($VariablesToExclude -contains $_.Name) })
						Write-Verbose "Found variables to import: $(($UserVariables | Select -expandproperty Name | Sort) -join ", " | Out-String).`n"
						
					}
					
					if ($ImportModules)
					{
						$UserModules = @(Get-Module | Where { $StandardUserEnv.Modules -notcontains $_.Name -and (Test-Path $_.Path -ErrorAction SilentlyContinue) } | Select -ExpandProperty Path)
						$UserSnapins = @(Get-PSSnapin | Select -ExpandProperty Name | Where { $StandardUserEnv.Snapins -notcontains $_ })
					}
				}
				
				#region Functions
				
				Function Get-RunspaceData
				{
					[cmdletbinding()]
					param ([switch]$Wait)
					
					#loop through runspaces
					#if $wait is specified, keep looping until all complete
					Do
					{
						
						#set more to false for tracking completion
						$more = $false
						
						#Progress bar if we have inputobject count (bound parameter)
						if (-not $Quiet)
						{
							Write-Progress -Activity "Running Query" -Status "Starting threads"`
										   -CurrentOperation "$startedCount threads defined - $totalCount input objects - $script:completedCount input objects processed"`
										   -PercentComplete $(Try { $script:completedCount / $totalCount * 100 }
							Catch { 0 })
						}
						
						#run through each runspace.           
						Foreach ($runspace in $runspaces)
						{
							
							#get the duration - inaccurate
							$currentdate = Get-Date
							$runtime = $currentdate - $runspace.startTime
							$runMin = [math]::Round($runtime.totalminutes, 2)
							
							#set up log object
							$log = "" | select Date, Action, Runtime, Status, Details
							$log.Action = "Removing:'$($runspace.object)'"
							$log.Date = $currentdate
							$log.Runtime = "$runMin minutes"
							
							#If runspace completed, end invoke, dispose, recycle, counter++
							If ($runspace.Runspace.isCompleted)
							{
								
								$script:completedCount++
								
								#check if there were errors
								if ($runspace.powershell.Streams.Error.Count -gt 0)
								{
									
									#set the logging info and move the file to completed
									$log.status = "CompletedWithErrors"
									Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
									foreach ($ErrorRecord in $runspace.powershell.Streams.Error)
									{
										Write-Error -ErrorRecord $ErrorRecord
									}
								}
								else
								{
									
									#add logging details and cleanup
									$log.status = "Completed"
									Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
								}
								
								#everything is logged, clean up the runspace
								$runspace.powershell.EndInvoke($runspace.Runspace)
								$runspace.powershell.dispose()
								$runspace.Runspace = $null
								$runspace.powershell = $null
								
							}
							
							#If runtime exceeds max, dispose the runspace
							ElseIf ($runspaceTimeout -ne 0 -and $runtime.totalseconds -gt $runspaceTimeout)
							{
								
								$script:completedCount++
								$timedOutTasks = $true
								
								#add logging details and cleanup
								$log.status = "TimedOut"
								Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
								Write-Error "Runspace timed out at $($runtime.totalseconds) seconds for the object:`n$($runspace.object | out-string)"
								
								#Depending on how it hangs, we could still get stuck here as dispose calls a synchronous method on the powershell instance
								if (!$noCloseOnTimeout) { $runspace.powershell.dispose() }
								$runspace.Runspace = $null
								$runspace.powershell = $null
								$completedCount++
								
							}
							
							#If runspace isn't null set more to true  
							ElseIf ($runspace.Runspace -ne $null)
							{
								$log = $null
								$more = $true
							}
							
							#log the results if a log file was indicated
							if ($logFile -and $log)
							{
								($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1] | out-file $LogFile -append
							}
						}
						
						#Clean out unused runspace jobs
						$temphash = $runspaces.clone()
						$temphash | Where { $_.runspace -eq $Null } | ForEach {
							$Runspaces.remove($_)
						}
						
						#sleep for a bit if we will loop again
						if ($PSBoundParameters['Wait']) { Start-Sleep -milliseconds $SleepTimer }
						
						#Loop again only if -wait parameter and there are more runspaces to process
					}
					while ($more -and $PSBoundParameters['Wait'])
					
					#End of runspace Function
				}
				
				#endregion Functions
				
				#region Init
				
				if ($PSCmdlet.ParameterSetName -eq 'ScriptFile')
				{
					$ScriptBlock = [scriptblock]::Create($(Get-Content $ScriptFile | out-string))
				}
				elseif ($PSCmdlet.ParameterSetName -eq 'ScriptBlock')
				{
					#Start building parameter names for the param block
					[string[]]$ParamsToAdd = '$_'
					if ($PSBoundParameters.ContainsKey('Parameter'))
					{
						$ParamsToAdd += '$Parameter'
					}
					
					$UsingVariableData = $Null
					
					
					# This code enables $Using support through the AST.
					# This is entirely from  Boe Prox, and his https://github.com/proxb/PoshRSJob module; all credit to Boe!
					
					if ($PSVersionTable.PSVersion.Major -gt 2)
					{
						#Extract using references
						$UsingVariables = $ScriptBlock.ast.FindAll({ $args[0] -is [System.Management.Automation.Language.UsingExpressionAst] }, $True)
						
						If ($UsingVariables)
						{
							$List = New-Object 'System.Collections.Generic.List`1[System.Management.Automation.Language.VariableExpressionAst]'
							ForEach ($Ast in $UsingVariables)
							{
								[void]$list.Add($Ast.SubExpression)
							}
							
							$UsingVar = $UsingVariables | Group Parent | ForEach { $_.Group | Select -First 1 }
							
							#Extract the name, value, and create replacements for each
							$UsingVariableData = ForEach ($Var in $UsingVar)
							{
								Try
								{
									$Value = Get-Variable -Name $Var.SubExpression.VariablePath.UserPath -ErrorAction Stop
									$NewName = ('$__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
									[pscustomobject]@{
										Name = $Var.SubExpression.Extent.Text
										Value = $Value.Value
										NewName = $NewName
										NewVarName = ('__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
									}
									$ParamsToAdd += $NewName
								}
								Catch
								{
									Write-Error "$($Var.SubExpression.Extent.Text) is not a valid Using: variable!"
								}
							}
							
							$NewParams = $UsingVariableData.NewName -join ', '
							$Tuple = [Tuple]::Create($list, $NewParams)
							$bindingFlags = [Reflection.BindingFlags]"Default,NonPublic,Instance"
							$GetWithInputHandlingForInvokeCommandImpl = ($ScriptBlock.ast.gettype().GetMethod('GetWithInputHandlingForInvokeCommandImpl', $bindingFlags))
							
							$StringScriptBlock = $GetWithInputHandlingForInvokeCommandImpl.Invoke($ScriptBlock.ast, @($Tuple))
							
							$ScriptBlock = [scriptblock]::Create($StringScriptBlock)
							
							Write-Verbose $StringScriptBlock
						}
					}
					
					$ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock("param($($ParamsToAdd -Join ", "))`r`n" + $Scriptblock.ToString())
				}
				else
				{
					Throw "Must provide ScriptBlock or ScriptFile"; Break
				}
				
				Write-Debug "`$ScriptBlock: $($ScriptBlock | Out-String)"
				Write-Verbose "Creating runspace pool and session states"
				
				#If specified, add variables and modules/snapins to session state
				$sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
				if ($ImportVariables)
				{
					if ($UserVariables.count -gt 0)
					{
						foreach ($Variable in $UserVariables)
						{
							$sessionstate.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $Variable.Name, $Variable.Value, $null))
						}
					}
				}
				if ($ImportModules)
				{
					if ($UserModules.count -gt 0)
					{
						foreach ($ModulePath in $UserModules)
						{
							$sessionstate.ImportPSModule($ModulePath)
						}
					}
					if ($UserSnapins.count -gt 0)
					{
						foreach ($PSSnapin in $UserSnapins)
						{
							[void]$sessionstate.ImportPSSnapIn($PSSnapin, [ref]$null)
						}
					}
				}
				
				#Create runspace pool
				$runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
				$runspacepool.Open()
				
				Write-Verbose "Creating empty collection to hold runspace jobs"
				$Script:runspaces = New-Object System.Collections.ArrayList
				
				#If inputObject is bound get a total count and set bound to true
				$global:__bound = $false
				$allObjects = @()
				if ($PSBoundParameters.ContainsKey("inputObject"))
				{
					$global:__bound = $true
				}
				
				#Set up log file if specified
				if ($LogFile)
				{
					New-Item -ItemType file -path $logFile -force | Out-Null
					("" | Select Date, Action, Runtime, Status, Details | ConvertTo-Csv -NoTypeInformation -Delimiter ";")[0] | Out-File $LogFile
				}
				
				#write initial log entry
				$log = "" | Select Date, Action, Runtime, Status, Details
				$log.Date = Get-Date
				$log.Action = "Batch processing started"
				$log.Runtime = $null
				$log.Status = "Started"
				$log.Details = $null
				if ($logFile)
				{
					($log | convertto-csv -Delimiter ";" -NoTypeInformation)[1] | Out-File $LogFile -Append
				}
				
				$timedOutTasks = $false
				
				#endregion INIT
			}
			
			Process
			{
				
				#add piped objects to all objects or set all objects to bound input object parameter
				if (-not $global:__bound)
				{
					$allObjects += $inputObject
				}
				else
				{
					$allObjects = $InputObject
				}
			}
			
			End
			{
				
				#Use Try/Finally to catch Ctrl+C and clean up.
				Try
				{
					#counts for progress
					$totalCount = $allObjects.count
					$script:completedCount = 0
					$startedCount = 0
					
					foreach ($object in $allObjects)
					{
						
						#region add scripts to runspace pool
						
						#Create the powershell instance, set verbose if needed, supply the scriptblock and parameters
						$powershell = [powershell]::Create()
						
						if ($VerbosePreference -eq 'Continue')
						{
							[void]$PowerShell.AddScript({ $VerbosePreference = 'Continue' })
						}
						
						[void]$PowerShell.AddScript($ScriptBlock).AddArgument($object)
						
						if ($parameter)
						{
							[void]$PowerShell.AddArgument($parameter)
						}
						
						# $Using support from Boe Prox
						if ($UsingVariableData)
						{
							Foreach ($UsingVariable in $UsingVariableData)
							{
								Write-Verbose "Adding $($UsingVariable.Name) with value: $($UsingVariable.Value)"
								[void]$PowerShell.AddArgument($UsingVariable.Value)
							}
						}
						
						#Add the runspace into the powershell instance
						$powershell.RunspacePool = $runspacepool
						
						#Create a temporary collection for each runspace
						$temp = "" | Select-Object PowerShell, StartTime, object, Runspace
						$temp.PowerShell = $powershell
						$temp.StartTime = Get-Date
						$temp.object = $object
						
						#Save the handle output when calling BeginInvoke() that will be used later to end the runspace
						$temp.Runspace = $powershell.BeginInvoke()
						$startedCount++
						
						#Add the temp tracking info to $runspaces collection
						Write-Verbose ("Adding {0} to collection at {1}" -f $temp.object, $temp.starttime.tostring())
						$runspaces.Add($temp) | Out-Null
						
						#loop through existing runspaces one time
						Get-RunspaceData
						
						#If we have more running than max queue (used to control timeout accuracy)
						#Script scope resolves odd PowerShell 2 issue
						$firstRun = $true
						while ($runspaces.count -ge $Script:MaxQueue)
						{
							
							#give verbose output
							if ($firstRun)
							{
								Write-Verbose "$($runspaces.count) items running - exceeded $Script:MaxQueue limit."
							}
							$firstRun = $false
							
							#run get-runspace data and sleep for a short while
							Get-RunspaceData
							Start-Sleep -Milliseconds $sleepTimer
							
						}
						
						#endregion add scripts to runspace pool
					}
					
					Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f (@($runspaces | Where { $_.Runspace -ne $Null }).Count))
					Get-RunspaceData -wait
					
					if (-not $quiet)
					{
						Write-Progress -Activity "Running Query" -Status "Starting threads" -Completed
					}
					
				}
				Finally
				{
					#Close the runspace pool, unless we specified no close on timeout and something timed out
					if (($timedOutTasks -eq $false) -or (($timedOutTasks -eq $true) -and ($noCloseOnTimeout -eq $false)))
					{
						Write-Verbose "Closing the runspace pool"
						$runspacepool.close()
					}
					
					#collect garbage
					[gc]::Collect()
				}
			}
		}
		
		Write-Verbose "PSBoundParameters = $($PSBoundParameters | Out-String)"
		
		$bound = $PSBoundParameters.keys -contains "ComputerName"
		if (-not $bound)
		{
			[System.Collections.ArrayList]$AllComputers = @()
		}
	}
	Process
	{
		
		#Handle both pipeline and bound parameter.  We don't want to stream objects, defeats purpose of parallelizing work
		if ($bound)
		{
			$AllComputers = $ComputerName
		}
		Else
		{
			foreach ($Computer in $ComputerName)
			{
				$AllComputers.add($Computer) | Out-Null
			}
		}
		
	}
	End
	{
		
		#Built up the parameters and run everything in parallel
		$params = @($Detail, $Quiet)
		$splat = @{
			Throttle = $Throttle
			RunspaceTimeout = $Timeout
			InputObject = $AllComputers
			parameter = $params
		}
		if ($NoCloseOnTimeout)
		{
			$splat.add('NoCloseOnTimeout', $True)
		}
		
		Invoke-Parallel @splat -ScriptBlock {
			
			$computer = $_.trim()
			$detail = $parameter[0]
			$quiet = $parameter[1]
			
			#They want detail, define and run test-server
			if ($detail)
			{
				Try
				{
					#Modification of jrich's Test-Server Function: https://gallery.technet.microsoft.com/scriptcenter/Powershell-Test-Server-e0cdea9a
					Function Test-Server
					{
						[cmdletBinding()]
						param (
							[parameter(
									   Mandatory = $true,
									   ValueFromPipeline = $true)]
							[string[]]$ComputerName,
							
							[switch]$All,
							
							[parameter(Mandatory = $false)]
							[switch]$CredSSP,
							
							[switch]$RemoteReg,
							
							[switch]$RDP,
							
							[switch]$RPC,
							
							[switch]$SMB,
							
							[switch]$WSMAN,
							
							[switch]$IPV6,
							
							[Management.Automation.PSCredential]$Credential
						)
						begin
						{
							$total = Get-Date
							$results = @()
							if ($credssp -and -not $Credential)
							{
								Throw "Must supply Credentials with CredSSP test"
							}
							
							[string[]]$props = write-output Name, IP, Domain, Ping, WSMAN, CredSSP, RemoteReg, RPC, RDP, SMB
							
							#Hash table to create PSObjects later, compatible with ps2...
							$Hash = @{ }
							foreach ($prop in $props)
							{
								$Hash.Add($prop, $null)
							}
							
							Function Test-Port
							{
								[cmdletbinding()]
								Param (
									[string]$srv,
									
									$port = 135,
									
									$timeout = 3000
								)
								$ErrorActionPreference = "SilentlyContinue"
								$tcpclient = new-Object system.Net.Sockets.TcpClient
								$iar = $tcpclient.BeginConnect($srv, $port, $null, $null)
								$wait = $iar.AsyncWaitHandle.WaitOne($timeout, $false)
								if (-not $wait)
								{
									$tcpclient.Close()
									Write-Verbose "Connection Timeout to $srv`:$port"
									$false
								}
								else
								{
									Try
									{
										$tcpclient.EndConnect($iar) | out-Null
										$true
									}
									Catch
									{
										write-verbose "Error for $srv`:$port`: $_"
										$false
									}
									$tcpclient.Close()
								}
							}
						}
						
						process
						{
							foreach ($name in $computername)
							{
								$dt = $cdt = Get-Date
								Write-verbose "Testing: $Name"
								$failed = 0
								try
								{
									$DNSEntity = [Net.Dns]::GetHostEntry($name)
									$domain = ($DNSEntity.hostname).replace("$name.", "")
									$ips = $DNSEntity.AddressList | %{
										if (-not (-not $IPV6 -and $_.AddressFamily -like "InterNetworkV6"))
										{
											$_.IPAddressToString
										}
									}
								}
								catch
								{
									$rst = New-Object -TypeName PSObject -Property $Hash | Select -Property $props
									$rst.name = $name
									$results += $rst
									$failed = 1
								}
								Write-verbose "DNS:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
								if ($failed -eq 0)
								{
									foreach ($ip in $ips)
									{
										
										$rst = New-Object -TypeName PSObject -Property $Hash | Select -Property $props
										$rst.name = $name
										$rst.ip = $ip
										$rst.domain = $domain
										
										if ($RDP -or $All)
										{
											####RDP Check (firewall may block rest so do before ping
											try
											{
												$socket = New-Object Net.Sockets.TcpClient($name, 3389) -ErrorAction stop
												if ($socket -eq $null)
												{
													$rst.RDP = $false
												}
												else
												{
													$rst.RDP = $true
													$socket.close()
												}
											}
											catch
											{
												$rst.RDP = $false
												Write-Verbose "Error testing RDP: $_"
											}
										}
										Write-verbose "RDP:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
										#########ping
										if (test-connection $ip -count 2 -Quiet)
										{
											Write-verbose "PING:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
											$rst.ping = $true
											
											if ($WSMAN -or $All)
											{
												try
												{
													############wsman
														Test-WSMan $ip -ErrorAction stop | Out-Null
														$rst.WSMAN = $true
													}
													catch
													{
														$rst.WSMAN = $false
														Write-Verbose "Error testing WSMAN: $_"
													}
													Write-verbose "WSMAN:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													if ($rst.WSMAN -and $credssp) ########### credssp
													{
														try
														{
															Test-WSMan $ip -Authentication Credssp -Credential $cred -ErrorAction stop
															$rst.CredSSP = $true
														}
														catch
														{
															$rst.CredSSP = $false
															Write-Verbose "Error testing CredSSP: $_"
														}
														Write-verbose "CredSSP:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													}
												}
												if ($RemoteReg -or $All)
												{
													try ########remote reg
													{
														[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ip) | Out-Null
														$rst.remotereg = $true
													}
													catch
													{
														$rst.remotereg = $false
														Write-Verbose "Error testing RemoteRegistry: $_"
													}
													Write-verbose "remote reg:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
												}
												if ($RPC -or $All)
												{
													try ######### wmi
													{
														$w = [wmi] ''
														$w.psbase.options.timeout = 15000000
														$w.path = "\\$Name\root\cimv2:Win32_ComputerSystem.Name='$Name'"
														$w | select none | Out-Null
														$rst.RPC = $true
													}
													catch
													{
														$rst.rpc = $false
														Write-Verbose "Error testing WMI/RPC: $_"
													}
													Write-verbose "WMI/RPC:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
												}
												if ($SMB -or $All)
												{
													
													#Use set location and resulting errors.  push and pop current location
													try ######### C$
													{
														$path = "\\$name\c$"
														Push-Location -Path $path -ErrorAction stop
														$rst.SMB = $true
														Pop-Location
													}
													catch
													{
														$rst.SMB = $false
														Write-Verbose "Error testing SMB: $_"
													}
													Write-verbose "SMB:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													
												}
											}
											else
											{
												$rst.ping = $false
												$rst.wsman = $false
												$rst.credssp = $false
												$rst.remotereg = $false
												$rst.rpc = $false
												$rst.smb = $false
											}
											$results += $rst
										}
									}
									Write-Verbose "Time for $($Name): $((New-TimeSpan $cdt ($dt)).totalseconds)"
									Write-Verbose "----------------------------"
								}
							}
							end
							{
								Write-Verbose "Time for all: $((New-TimeSpan $total ($dt)).totalseconds)"
								Write-Verbose "----------------------------"
								return $results
							}
						}
						
						#Build up parameters for Test-Server and run it
						$TestServerParams = @{
							ComputerName = $Computer
							ErrorAction = "Stop"
						}
						
						if ($detail -eq "*")
						{
							$detail = "WSMan", "RemoteReg", "RPC", "RDP", "SMB"
						}
						
						$detail | Select -Unique | Foreach-Object { $TestServerParams.add($_, $True) }
						Test-Server @TestServerParams | Select -Property $("Name", "IP", "Domain", "Ping" + $detail)
					}
					Catch
					{
						Write-Warning "Error with Test-Server: $_"
					}
				}
				#We just want ping output
				else
				{
					Try
					{
						#Pick out a few properties, add a status label.  If quiet output, just return the address
						$result = $null
						if ($result = @(Test-Connection -ComputerName $computer -Count 2 -erroraction Stop))
						{
							$Output = $result | Select -first 1 -Property Address,
													   IPV4Address,
													   IPV6Address,
													   ResponseTime,
													   @{ label = "STATUS"; expression = { "Responding" } }
							
							if ($quiet)
							{
								$Output.address
							}
							else
							{
								$Output
							}
						}
					}
					Catch
					{
						if (-not $quiet)
						{
							#Ping failed.  I'm likely making inappropriate assumptions here, let me know if this is the case : )
							if ($_ -match "No such host is known")
							{
								$status = "Unknown host"
							}
							elseif ($_ -match "Error due to lack of resources")
							{
								$status = "No Response"
							}
							else
							{
								$status = "Error: $_"
							}
							
							"" | Select -Property @{ label = "Address"; expression = { $computer } },
										IPV4Address,
										IPV6Address,
										ResponseTime,
										@{ label = "STATUS"; expression = { $status } }
						}
					}
				}
			}
		}
	}
## End Invoke-Ping
## Begin WhoIs
Function WhoIs {
<#
.SYNOPSIS
Domain name WhoIs
.DESCRIPTION
Performs a domain name lookup and returns information such as
domain availability (creation and expiration date),
domain ownership, name servers, etc..

.PARAMETER domain
Specifies the domain name (enter the domain name without http:// and www (e.g. power-shell.com))

.EXAMPLE
WhoIs -domain power-shell.com 
whois power-shell.com

.NOTES
File Name: whois.ps1
Author: Nikolay Petkov
Blog: http://power-shell.com
Last Edit: 12/20/2014

.LINK
http://power-shell.com
#>
param (
                [Parameter(Mandatory=$True,
                           HelpMessage='Please enter domain name (e.g. microsoft.com)')]
                           [string]$domain
        )
Write-Host "Connecting to Web Services URL..." -ForegroundColor Green
try {
#Retrieve the data from web service WSDL
If ($whois = New-WebServiceProxy -uri "http://www.webservicex.net/whois.asmx?WSDL") {Write-Host "Ok" -ForegroundColor Green}
else {Write-Host "Error" -ForegroundColor Red}
Write-Host "Gathering $domain data..." -ForegroundColor Green
#Return the data
(($whois.getwhois("=$domain")).Split("<<<")[0])
} catch {
## End WhoIs
Write-Host "Please enter valid domain name (e.g. microsoft.com)." -ForegroundColor Red}
}
## End WhoIs
## Begin Get-NetworkStatistics
Function Get-NetworkStatistics{
<#
.SYNOPSIS
PowerShell version of netstat
.EXAMPLE
Get-NetworkStatistics
.EXAMPLE
Get-NetworkStatistics | where-object {$_.State -eq "LISTENING"} | Format-Table
#>	
    $properties = 'Protocol','LocalAddress','LocalPort' 
    $properties += 'RemoteAddress','RemotePort','State','ProcessName','PID' 

    netstat -ano | Select-String -Pattern '\s+(TCP|UDP)' | ForEach-Object { 

        $item = $_.line.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries) 

        if($item[1] -notmatch '^\[::') 
        {            
            if (($la = $item[1] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6') 
            { 
               $localAddress = $la.IPAddressToString 
               $localPort = $item[1].split('\]:')[-1] 
            } 
            else 
            { 
                $localAddress = $item[1].split(':')[0] 
                $localPort = $item[1].split(':')[-1] 
            }  

            if (($ra = $item[2] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6') 
            { 
               $remoteAddress = $ra.IPAddressToString 
               $remotePort = $item[2].split('\]:')[-1] 
            } 
            else 
            { 
               $remoteAddress = $item[2].split(':')[0] 
               $remotePort = $item[2].split(':')[-1] 
            }  

            New-Object PSObject -Property @{ 
                PID = $item[-1] 
                ProcessName = (Get-Process -Id $item[-1] -ErrorAction SilentlyContinue).Name 
                Protocol = $item[0] 
                LocalAddress = $localAddress 
                LocalPort = $localPort 
                RemoteAddress =$remoteAddress 
                RemotePort = $remotePort 
                State = if($item[0] -eq 'tcp') {$item[3]} else {$null} 
            } | Select-Object -Property $properties 
        } 
    } 
}
## End Get-NetworkStatistics
## Begin Get-NetworkInfo
Function Get-NetworkInfo {
    <#   
        .SYNOPSIS   
            Retrieves the network configuration from a local or remote client.      
             
        .DESCRIPTION   
            Retrieves the network configuration from a local or remote client.        
        
        .PARAMETER Computername
            A single or collection of systems to perform the query against
        
        .PARAMETER Credential
            Alternate credentials to use for query of network information        
        
        .PARAMETER Throttle
            Number of asynchonous jobs that will run at a time
        
        .NOTES   
            Name: Get-NetworkInfo.ps1
            Author: Boe Prox
            Version: 1.0
        
        .EXAMPLE 
             Get-NetworkInfo -Computername 'System1'
            
            NICDescription : Ethernet Network Adapter
            MACAddress     : 00:11:22:33:aa:bb
            NICName        : enthad
            Computername   : System1.domain.com
            DHCPEnabled    : True
            WINSPrimary    : 192.0.0.25
            SubnetMask     : {255.255.255.255}
            WINSSecondary  : 192.0.0.26
            DNSServer      : {192.0.0.31, 192.0.0.30}
            IPAddress      : {192.0.0.5}
            DefaultGateway : {192.0.0.1}         
             
            Description 
            ----------- 
            Retrieves the network information from 'System1'      

        .EXAMPLE
            $Servers = Get-Content Servers.txt
            $Servers | Get-NetworkInfo -Throttle 10
            
            Description
            -----------
            Retrieves all of network information from the remote servers while running 10 runspace jobs at a time.  
            
        .EXAMPLE
            (Get-Content Servers.txt) | Get-NetworkInfo -Credential domain\adminuser -Throttle 10
            
            Description
            -----------
            Gathers all of the network information from the systems in the text file. Also uses alternate administrator credentials provided.                                            
    #>
    
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('CN','__Server','IPAddress','Server')]
        [string[]]$Computername = $Env:Computername,
        
        [parameter()]
        [Alias('RunAs')]
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty,       
        
        [parameter()]
        [int]$Throttle = 15
    )
    Begin {
        #Function that will be used to process runspace jobs
        Function Get-RunspaceData {
            [cmdletbinding()]
            param(
                [switch]$Wait
            )
            Do {
                $more = $false         
                Foreach($runspace in $runspaces) {
                    If ($runspace.Runspace.isCompleted) {
                        $runspace.powershell.EndInvoke($runspace.Runspace)
                        $runspace.powershell.dispose()
                        $runspace.Runspace = $null
                        $runspace.powershell = $null
                        $Script:i++                  
                    } ElseIf ($runspace.Runspace -ne $null) {
                        $more = $true
                    }
                }
                If ($more -AND $PSBoundParameters['Wait']) {
                    Start-Sleep -Milliseconds 100
                }   
                #Clean out unused runspace jobs
                $temphash = $runspaces.clone()
                $temphash | Where {
                    $_.runspace -eq $Null
                } | ForEach {
                    Write-Verbose ("Removing {0}" -f $_.computer)
                    $Runspaces.remove($_)
                }             
            } while ($more -AND $PSBoundParameters['Wait'])
        }
            
        Write-Verbose ("Performing inital Administrator check")
        $usercontext = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        $IsAdmin = $usercontext.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")                   
        
        #Main collection to hold all data returned from runspace jobs
        $Script:report = @()    
        
        Write-Verbose ("Building hash table for WMI parameters")
        $WMIhash = @{
            Class = "Win32_NetworkAdapterConfiguration"
            Filter = "IPEnabled='$True'"
            ErrorAction = "Stop"
        } 
        
        #Supplied Alternate Credentials?
        If ($PSBoundParameters['Credential']) {
            $wmihash.credential = $Credential
        }
        
        #Define hash table for Get-RunspaceData Function
        $runspacehash = @{}

        #Define Scriptblock for runspaces
        $scriptblock = {
            Param (
                $Computer,
                $wmihash
            )           
            Write-Verbose ("{0}: Checking network connection" -f $Computer)
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                #Check if running against local system and perform necessary actions
                Write-Verbose ("Checking for local system")
                If ($Computer -eq $Env:Computername) {
                    $wmihash.remove('Credential')
                } Else {
                    $wmihash.Computername = $Computer
                }
                Try {
                        Get-WmiObject @WMIhash | ForEach {
                            $IpHash =  @{
                                Computername = $_.DNSHostName
                                DNSDomain = $_.DNSDomain
                                IPAddress = $_.IpAddress
                                SubnetMask = $_.IPSubnet
                                DefaultGateway = $_.DefaultIPGateway
                                DNSServer = $_.DNSServerSearchOrder
                                DHCPEnabled = $_.DHCPEnabled
                                MACAddress  = $_.MACAddress
                                WINSPrimary = $_.WINSPrimaryServer
                                WINSSecondary = $_.WINSSecondaryServer
                                NICName = $_.ServiceName
                                NICDescription = $_.Description
                            }
                            $IpStack = New-Object PSObject -Property $IpHash
                            #Add a unique object typename
                            $IpStack.PSTypeNames.Insert(0,"IPStack.Information")
                            $IpStack 
                        }
                    } Catch {
                        Write-Warning ("{0}: {1}" -f $Computer,$_.Exception.Message)
                        Break
                }
            } Else {
                Write-Warning ("{0}: Unavailable!" -f $Computer)
                Break
            }        
        }
        
        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
        $runspacepool.Open()  
        
        Write-Verbose ("Creating empty collection to hold runspace jobs")
        $Script:runspaces = New-Object System.Collections.ArrayList        
    }
    Process {        
        $totalcount = $computername.count
        Write-Verbose ("Validating that current user is Administrator or supplied alternate credentials")        
        If (-Not ($Computername.count -eq 1 -AND $Computername[0] -eq $Env:Computername)) {
            #Now check that user is either an Administrator or supplied Alternate Credentials
            If (-Not ($IsAdmin -OR $PSBoundParameters['Credential'])) {
                Write-Warning ("You must be an Administrator to perform this action against remote systems!")
                Break
            }
        }
        ForEach ($Computer in $Computername) {
           #Create the powershell instance and supply the scriptblock with the other parameters 
           $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($wmihash)
           
           #Add the runspace into the powershell instance
           $powershell.RunspacePool = $runspacepool
           
           #Create a temporary collection for each runspace
           $temp = "" | Select-Object PowerShell,Runspace,Computer
           $Temp.Computer = $Computer
           $temp.PowerShell = $powershell
           
           #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
           $temp.Runspace = $powershell.BeginInvoke()
           Write-Verbose ("Adding {0} collection" -f $temp.Computer)
           $runspaces.Add($temp) | Out-Null
           
           Write-Verbose ("Checking status of runspace jobs")
           Get-RunspaceData @runspacehash
        }                        
    }
    End {                     
        Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f (@(($runspaces | Where {$_.Runspace -ne $Null}).Count)))
        $runspacehash.Wait = $true
        Get-RunspaceData @runspacehash
        
        Write-Verbose ("Closing the runspace pool")
        $runspacepool.close()               
    }
}
## End Get-NetworkInfo
## Begin Get-SnmpTrap
Function Get-SnmpTrap {
<#
.SYNOPSIS
## Begin that
Function that will list SNMP Community string, Security Options and Trap Configuration for SNMP version 1 and version 2c.
.DESCRIPTION
** This Function will list SNMP settings of  windows server by reading the registry keys under HKLM\SYSTEM\CurrentControlSet\services\SNMP\Parameters\
Example usage:																					  
Get_SnmpTrap
This will list the  SNMP Community string, Security Options and Trap Configuration on the server. The meaning of each column is:
AcceptedCommunityStrings => The community string that the SNMP agent is allowed to receive. If the host is not requested with one of these pre-defined 
community strings, then the host will send an authentication trap.
AllowedHosts => The hostnames or IP addresses from which SNMP agent will accept SNMP messages.
CommunityRights => The permission that determines how the SNMP agent processes the incoming request from various communities.
TrapCommunityNames => When an SNMP agent receives a request that does not contain a valid community name or the host that is sending the message 
is not on the list of acceptable hosts, the agent can send an authentication trap message to one or more trap destinations (management systems)
TrapDestinations => The host names or IP addresses of trap destinations which are defined under the TrapCommunityNames.
SendTrap => It indicates whether sending autentication trap is enabled.
Author: phyoepaing3.142@gmail.com
Country: Myanmar(Burma)
Released: 05/07/2017
.EXAMPLE
Get_SnmpTrap
This will list the  SNMP Community string, Security Options and Trap Configuration on the server.
.LINK
You can find this script and more at: https://www.sysadminplus.blogspot.com/
#>

### DATA lookup section to convert registry numeric to corresponding output ###
$ConvertRights = DATA { ConvertFrom-StringData -StringData @'
1 = NONE
2 = NOTIFY
4 = READ-ONLY
8 = READ-WRITE
16 = READ-CREATE
'@}

$rh = '2147483650';  ## This number represents HKLM
$key1 = 'SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers';
$reg = [wmiclass]"\\localhost\root\default:StdRegprov"; 
$obj = New-Object -TypeName PsObject -Property @{AllowedHosts=@(); AcceptedCommunityStrings="";  CommunityRights =@(); TrapCommunityNames=@(); TrapDestinations=@(); SendTrap="" }; 
$AccessDenied = 0;

### Read the registry to find the allowed hosts for incoming community string ###
$i=1;
while ( $reg.GetStringValue($rh, $key1, $i ).sValue )
	{
	$obj.AllowedHosts += $reg.GetStringValue($rh, $key1, $i ).sValue; 
	$i ++;
	}
If ($obj.AllowedHosts.count -eq 1)
	{
	$obj.AllowedHosts = $obj.AllowedHosts[0];
	}

### Read the Community Strings ###	
Try {
	$obj.AcceptedCommunityStrings = (Gi -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities -EA Stop).Property; 

## If there is only one community string, then convert the property type to string from array ###
	If ($obj.AcceptedCommunityStrings.count -eq 1)
	{ $obj.AcceptedCommunityStrings = $obj.AcceptedCommunityStrings[0] }

### If there are multiple community strings, then read through all the security permission of each community string 	via registry ##
	If ($obj.AcceptedCommunityStrings -is [array])
	{	
		$obj.AcceptedCommunityStrings | foreach {
		$securityRight =  [string]((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities).$_)
		$obj.CommunityRights += $_+":"+$ConvertRights[$securityRight]
			}
		}
	else
		{
		[string]$securityRight = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities).$($obj.AcceptedCommunityStrings)
		$obj.CommunityRights = $ConvertRights[$securityRight]
		}
	}
catch [System.Security.SecurityException]
	{ 
	Write-Host -fore red "Access to Registry is denied. Please make sure you have permission to access registry and run in the elevated command prompt.`n"; 
	$obj.AllowedHosts = "N/A"
	$obj.AcceptedCommunityStrings = "N/A"
	$obj.CommunityRights = "N/A"
	$obj.TrapCommunityNames = "N/A"
	$obj.TrapDestinations = "N/A"
	$obj.SendTrap = "N/A"
	$AccessDenied = 1; 
	}
catch 
	{ 
	Write-Host -fore red "SMNP Service is not installed on one or more servers.`n"; 
	$obj.AllowedHosts = "N/A"
	$obj.AcceptedCommunityStrings = "N/A"
	$obj.CommunityRights = "N/A"
	$obj.TrapCommunityNames = "N/A"
	$obj.TrapDestinations = "N/A"
	$obj.SendTrap = "N/A"
	$AccessDenied = 1; 
	$obj;
	}

## If the read of registry is not access-denied from previous try-catch statement, then continue ##	
If (!$AccessDenied)	
	{
	Try {
		$TrapConfig = Gci -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration -EA Stop ;
		$TrapConfig | foreach {
			$obj.TrapCommunityNames += $_.PsChildName
			}
	If ($obj.TrapCommunityNames.count -eq 1)	
	{ $obj.TrapCommunityNames = $obj.TrapCommunityNames[0]	}
			
		}
	catch 
		{  }
		
### Find destination for each Trap. The trap's community name will be prefixed on the trap's destination IP/hosts if there are multiple Traps configured, if it's single trap, then use without prefix ###

If ($obj.TrapCommunityNames -is [array])
	{
		$obj.TrapCommunityNames | foreach {
		$key2 = "SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\$_";
		$i=1;
			while ( $reg.GetStringValue($rh, $key2, $i ).sValue )
				{
				$obj.TrapDestinations += $_+":"+$reg.GetStringValue($rh, $key2, $i ).sValue; 
				$i ++;
				}
		}
	}
else
	{
	$key2 = "SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\$($obj.TrapCommunityNames)";
	$i=1;
		while ( $reg.GetStringValue($rh, $key2, $i ).sValue )
				{
				$obj.TrapDestinations += $reg.GetStringValue($rh, $key2, $i ).sValue; 
				$i ++;
				}
		}
	
### If there is only one entry in the Trap Destination, then convert  the array to string ###
If ($obj.TrapDestinations.count -eq 1)
	{
	$obj.TrapDestinations = $obj.TrapDestinations[0];
	}
	
#### Check if the 'Send Authentication Trap' check box is enabled ###
	Switch ((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters).EnableAuthenticationTraps)
		{
		"0" { $obj.SendTrap = "Disabled" }
		"1" { $obj.SendTrap = "Enabled "}		
		}
	$obj;	
	}
}
## End Get-SNMPTrap
## Begin Invoke-Ping
Function Invoke-Ping{
<#
.SYNOPSIS
    Ping or test connectivity to systems in parallel
    
.DESCRIPTION
    Ping or test connectivity to systems in parallel

    Default action will run a ping against systems
        If Quiet parameter is specified, we return an array of systems that responded
        If Detail parameter is specified, we test WSMan, RemoteReg, RPC, RDP and/or SMB

.PARAMETER ComputerName
    One or more computers to test

.PARAMETER Quiet
    If specified, only return addresses that responded to Test-Connection

.PARAMETER Detail
    Include one or more additional tests as specified:
        WSMan      via Test-WSMan
        RemoteReg  via Microsoft.Win32.RegistryKey
        RPC        via WMI
        RDP        via port 3389
        SMB        via \\ComputerName\C$
        *          All tests

.PARAMETER Timeout
    Time in seconds before we attempt to dispose an individual query.  Default is 20

.PARAMETER Throttle
    Throttle query to this many parallel runspaces.  Default is 100.

.PARAMETER NoCloseOnTimeout
    Do not dispose of timed out tasks or attempt to close the runspace if threads have timed out

    This will prevent the script from hanging in certain situations where threads become non-responsive, at the expense of leaking memory within the PowerShell host.

.EXAMPLE
    Invoke-Ping Server1, Server2, Server3 -Detail *

    # Check for WSMan, Remote Registry, Remote RPC, RDP, and SMB (via C$) connectivity against 3 machines

.EXAMPLE
    $Computers | Invoke-Ping

    # Ping computers in $Computers in parallel

.EXAMPLE
    $Responding = $Computers | Invoke-Ping -Quiet
    
    # Create a list of computers that successfully responded to Test-Connection

.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Invoke-Ping-Test-in-b553242a

.FunctionALITY
    Computers
	
.NOTES
	Warren F
	http://ramblingcookiemonster.github.io/Invoke-Ping/

#>
	[cmdletbinding(DefaultParameterSetName = 'Ping')]
	param (
		[Parameter(ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[string[]]$ComputerName,
		
		[Parameter(ParameterSetName = 'Detail')]
		[validateset("*", "WSMan", "RemoteReg", "RPC", "RDP", "SMB")]
		[string[]]$Detail,
		
		[Parameter(ParameterSetName = 'Ping')]
		[switch]$Quiet,
		
		[int]$Timeout = 20,
		
		[int]$Throttle = 100,
		
		[switch]$NoCloseOnTimeout
	)
	Begin
	{
		
		#http://gallery.technet.microsoft.com/Run-Parallel-Parallel-377fd430
		Function Invoke-Parallel
		{
			[cmdletbinding(DefaultParameterSetName = 'ScriptBlock')]
			Param (
				[Parameter(Mandatory = $false, position = 0, ParameterSetName = 'ScriptBlock')]
				[System.Management.Automation.ScriptBlock]$ScriptBlock,
				
				[Parameter(Mandatory = $false, ParameterSetName = 'ScriptFile')]
				[ValidateScript({ test-path $_ -pathtype leaf })]
				$ScriptFile,
				
				[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
				[Alias('CN', '__Server', 'IPAddress', 'Server', 'ComputerName')]
				[PSObject]$InputObject,
				
				[PSObject]$Parameter,
				
				[switch]$ImportVariables,
				
				[switch]$ImportModules,
				
				[int]$Throttle = 20,
				
				[int]$SleepTimer = 200,
				
				[int]$RunspaceTimeout = 0,
				
				[switch]$NoCloseOnTimeout = $false,
				
				[int]$MaxQueue,
				
				[validatescript({ Test-Path (Split-Path $_ -parent) })]
				[string]$LogFile = "C:\temp\log.log",
				
				[switch]$Quiet = $false
			)
			
			Begin
			{
				
				#No max queue specified?  Estimate one.
				#We use the script scope to resolve an odd PowerShell 2 issue where MaxQueue isn't seen later in the Function
				if (-not $PSBoundParameters.ContainsKey('MaxQueue'))
				{
					if ($RunspaceTimeout -ne 0) { $script:MaxQueue = $Throttle }
					else { $script:MaxQueue = $Throttle * 3 }
				}
				else
				{
					$script:MaxQueue = $MaxQueue
				}
				
				Write-Verbose "Throttle: '$throttle' SleepTimer '$sleepTimer' runSpaceTimeout '$runspaceTimeout' maxQueue '$maxQueue' logFile '$logFile'"
				
				#If they want to import variables or modules, create a clean runspace, get loaded items, use those to exclude items
				if ($ImportVariables -or $ImportModules)
				{
					$StandardUserEnv = [powershell]::Create().addscript({
						
						#Get modules and snapins in this clean runspace
						$Modules = Get-Module | Select -ExpandProperty Name
						$Snapins = Get-PSSnapin | Select -ExpandProperty Name
						
						#Get variables in this clean runspace
						#Called last to get vars like $? into session
						$Variables = Get-Variable | Select -ExpandProperty Name
						
						#Return a hashtable where we can access each.
						@{
							Variables = $Variables
							Modules = $Modules
							Snapins = $Snapins
						}
					}).invoke()[0]
					
					if ($ImportVariables)
					{
						#Exclude common parameters, bound parameters, and automatic variables
						Function _temp { [cmdletbinding()]
							param () }
						$VariablesToExclude = @((Get-Command _temp | Select -ExpandProperty parameters).Keys + $PSBoundParameters.Keys + $StandardUserEnv.Variables)
						Write-Verbose "Excluding variables $(($VariablesToExclude | sort) -join ", ")"
						
						# we don't use 'Get-Variable -Exclude', because it uses regexps. 
						# One of the veriables that we pass is '$?'. 
						# There could be other variables with such problems.
						# Scope 2 required if we move to a real module
						$UserVariables = @(Get-Variable | Where { -not ($VariablesToExclude -contains $_.Name) })
						Write-Verbose "Found variables to import: $(($UserVariables | Select -expandproperty Name | Sort) -join ", " | Out-String).`n"
						
					}
					
					if ($ImportModules)
					{
						$UserModules = @(Get-Module | Where { $StandardUserEnv.Modules -notcontains $_.Name -and (Test-Path $_.Path -ErrorAction SilentlyContinue) } | Select -ExpandProperty Path)
						$UserSnapins = @(Get-PSSnapin | Select -ExpandProperty Name | Where { $StandardUserEnv.Snapins -notcontains $_ })
					}
				}
				
				#region Functions
				
				Function Get-RunspaceData
				{
					[cmdletbinding()]
					param ([switch]$Wait)
					
					#loop through runspaces
					#if $wait is specified, keep looping until all complete
					Do
					{
						
						#set more to false for tracking completion
						$more = $false
						
						#Progress bar if we have inputobject count (bound parameter)
						if (-not $Quiet)
						{
							Write-Progress -Activity "Running Query" -Status "Starting threads"`
										   -CurrentOperation "$startedCount threads defined - $totalCount input objects - $script:completedCount input objects processed"`
										   -PercentComplete $(Try { $script:completedCount / $totalCount * 100 }
							Catch { 0 })
						}
						
						#run through each runspace.           
						Foreach ($runspace in $runspaces)
						{
							
							#get the duration - inaccurate
							$currentdate = Get-Date
							$runtime = $currentdate - $runspace.startTime
							$runMin = [math]::Round($runtime.totalminutes, 2)
							
							#set up log object
							$log = "" | select Date, Action, Runtime, Status, Details
							$log.Action = "Removing:'$($runspace.object)'"
							$log.Date = $currentdate
							$log.Runtime = "$runMin minutes"
							
							#If runspace completed, end invoke, dispose, recycle, counter++
							If ($runspace.Runspace.isCompleted)
							{
								
								$script:completedCount++
								
								#check if there were errors
								if ($runspace.powershell.Streams.Error.Count -gt 0)
								{
									
									#set the logging info and move the file to completed
									$log.status = "CompletedWithErrors"
									Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
									foreach ($ErrorRecord in $runspace.powershell.Streams.Error)
									{
										Write-Error -ErrorRecord $ErrorRecord
									}
								}
								else
								{
									
									#add logging details and cleanup
									$log.status = "Completed"
									Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
								}
								
								#everything is logged, clean up the runspace
								$runspace.powershell.EndInvoke($runspace.Runspace)
								$runspace.powershell.dispose()
								$runspace.Runspace = $null
								$runspace.powershell = $null
								
							}
							
							#If runtime exceeds max, dispose the runspace
							ElseIf ($runspaceTimeout -ne 0 -and $runtime.totalseconds -gt $runspaceTimeout)
							{
								
								$script:completedCount++
								$timedOutTasks = $true
								
								#add logging details and cleanup
								$log.status = "TimedOut"
								Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
								Write-Error "Runspace timed out at $($runtime.totalseconds) seconds for the object:`n$($runspace.object | out-string)"
								
								#Depending on how it hangs, we could still get stuck here as dispose calls a synchronous method on the powershell instance
								if (!$noCloseOnTimeout) { $runspace.powershell.dispose() }
								$runspace.Runspace = $null
								$runspace.powershell = $null
								$completedCount++
								
							}
							
							#If runspace isn't null set more to true  
							ElseIf ($runspace.Runspace -ne $null)
							{
								$log = $null
								$more = $true
							}
							
							#log the results if a log file was indicated
							if ($logFile -and $log)
							{
								($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1] | out-file $LogFile -append
							}
						}
						
						#Clean out unused runspace jobs
						$temphash = $runspaces.clone()
						$temphash | Where { $_.runspace -eq $Null } | ForEach {
							$Runspaces.remove($_)
						}
						
						#sleep for a bit if we will loop again
						if ($PSBoundParameters['Wait']) { Start-Sleep -milliseconds $SleepTimer }
						
						#Loop again only if -wait parameter and there are more runspaces to process
					}
					while ($more -and $PSBoundParameters['Wait'])
					
					#End of runspace Function
				}
				
				#endregion Functions
				
				#region Init
				
				if ($PSCmdlet.ParameterSetName -eq 'ScriptFile')
				{
					$ScriptBlock = [scriptblock]::Create($(Get-Content $ScriptFile | out-string))
				}
				elseif ($PSCmdlet.ParameterSetName -eq 'ScriptBlock')
				{
					#Start building parameter names for the param block
					[string[]]$ParamsToAdd = '$_'
					if ($PSBoundParameters.ContainsKey('Parameter'))
					{
						$ParamsToAdd += '$Parameter'
					}
					
					$UsingVariableData = $Null
					
					
					# This code enables $Using support through the AST.
					# This is entirely from  Boe Prox, and his https://github.com/proxb/PoshRSJob module; all credit to Boe!
					
					if ($PSVersionTable.PSVersion.Major -gt 2)
					{
						#Extract using references
						$UsingVariables = $ScriptBlock.ast.FindAll({ $args[0] -is [System.Management.Automation.Language.UsingExpressionAst] }, $True)
						
						If ($UsingVariables)
						{
							$List = New-Object 'System.Collections.Generic.List`1[System.Management.Automation.Language.VariableExpressionAst]'
							ForEach ($Ast in $UsingVariables)
							{
								[void]$list.Add($Ast.SubExpression)
							}
							
							$UsingVar = $UsingVariables | Group Parent | ForEach { $_.Group | Select -First 1 }
							
							#Extract the name, value, and create replacements for each
							$UsingVariableData = ForEach ($Var in $UsingVar)
							{
								Try
								{
									$Value = Get-Variable -Name $Var.SubExpression.VariablePath.UserPath -ErrorAction Stop
									$NewName = ('$__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
									[pscustomobject]@{
										Name = $Var.SubExpression.Extent.Text
										Value = $Value.Value
										NewName = $NewName
										NewVarName = ('__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
									}
									$ParamsToAdd += $NewName
								}
								Catch
								{
									Write-Error "$($Var.SubExpression.Extent.Text) is not a valid Using: variable!"
								}
							}
							
							$NewParams = $UsingVariableData.NewName -join ', '
							$Tuple = [Tuple]::Create($list, $NewParams)
							$bindingFlags = [Reflection.BindingFlags]"Default,NonPublic,Instance"
							$GetWithInputHandlingForInvokeCommandImpl = ($ScriptBlock.ast.gettype().GetMethod('GetWithInputHandlingForInvokeCommandImpl', $bindingFlags))
							
							$StringScriptBlock = $GetWithInputHandlingForInvokeCommandImpl.Invoke($ScriptBlock.ast, @($Tuple))
							
							$ScriptBlock = [scriptblock]::Create($StringScriptBlock)
							
							Write-Verbose $StringScriptBlock
						}
					}
					
					$ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock("param($($ParamsToAdd -Join ", "))`r`n" + $Scriptblock.ToString())
				}
				else
				{
					Throw "Must provide ScriptBlock or ScriptFile"; Break
				}
				
				Write-Debug "`$ScriptBlock: $($ScriptBlock | Out-String)"
				Write-Verbose "Creating runspace pool and session states"
				
				#If specified, add variables and modules/snapins to session state
				$sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
				if ($ImportVariables)
				{
					if ($UserVariables.count -gt 0)
					{
						foreach ($Variable in $UserVariables)
						{
							$sessionstate.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $Variable.Name, $Variable.Value, $null))
						}
					}
				}
				if ($ImportModules)
				{
					if ($UserModules.count -gt 0)
					{
						foreach ($ModulePath in $UserModules)
						{
							$sessionstate.ImportPSModule($ModulePath)
						}
					}
					if ($UserSnapins.count -gt 0)
					{
						foreach ($PSSnapin in $UserSnapins)
						{
							[void]$sessionstate.ImportPSSnapIn($PSSnapin, [ref]$null)
						}
					}
				}
				
				#Create runspace pool
				$runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
				$runspacepool.Open()
				
				Write-Verbose "Creating empty collection to hold runspace jobs"
				$Script:runspaces = New-Object System.Collections.ArrayList
				
				#If inputObject is bound get a total count and set bound to true
				$global:__bound = $false
				$allObjects = @()
				if ($PSBoundParameters.ContainsKey("inputObject"))
				{
					$global:__bound = $true
				}
				
				#Set up log file if specified
				if ($LogFile)
				{
					New-Item -ItemType file -path $logFile -force | Out-Null
					("" | Select Date, Action, Runtime, Status, Details | ConvertTo-Csv -NoTypeInformation -Delimiter ";")[0] | Out-File $LogFile
				}
				
				#write initial log entry
				$log = "" | Select Date, Action, Runtime, Status, Details
				$log.Date = Get-Date
				$log.Action = "Batch processing started"
				$log.Runtime = $null
				$log.Status = "Started"
				$log.Details = $null
				if ($logFile)
				{
					($log | convertto-csv -Delimiter ";" -NoTypeInformation)[1] | Out-File $LogFile -Append
				}
				
				$timedOutTasks = $false
				
				#endregion INIT
			}
			
			Process
			{
				
				#add piped objects to all objects or set all objects to bound input object parameter
				if (-not $global:__bound)
				{
					$allObjects += $inputObject
				}
				else
				{
					$allObjects = $InputObject
				}
			}
			
			End
			{
				
				#Use Try/Finally to catch Ctrl+C and clean up.
				Try
				{
					#counts for progress
					$totalCount = $allObjects.count
					$script:completedCount = 0
					$startedCount = 0
					
					foreach ($object in $allObjects)
					{
						
						#region add scripts to runspace pool
						
						#Create the powershell instance, set verbose if needed, supply the scriptblock and parameters
						$powershell = [powershell]::Create()
						
						if ($VerbosePreference -eq 'Continue')
						{
							[void]$PowerShell.AddScript({ $VerbosePreference = 'Continue' })
						}
						
						[void]$PowerShell.AddScript($ScriptBlock).AddArgument($object)
						
						if ($parameter)
						{
							[void]$PowerShell.AddArgument($parameter)
						}
						
						# $Using support from Boe Prox
						if ($UsingVariableData)
						{
							Foreach ($UsingVariable in $UsingVariableData)
							{
								Write-Verbose "Adding $($UsingVariable.Name) with value: $($UsingVariable.Value)"
								[void]$PowerShell.AddArgument($UsingVariable.Value)
							}
						}
						
						#Add the runspace into the powershell instance
						$powershell.RunspacePool = $runspacepool
						
						#Create a temporary collection for each runspace
						$temp = "" | Select-Object PowerShell, StartTime, object, Runspace
						$temp.PowerShell = $powershell
						$temp.StartTime = Get-Date
						$temp.object = $object
						
						#Save the handle output when calling BeginInvoke() that will be used later to end the runspace
						$temp.Runspace = $powershell.BeginInvoke()
						$startedCount++
						
						#Add the temp tracking info to $runspaces collection
						Write-Verbose ("Adding {0} to collection at {1}" -f $temp.object, $temp.starttime.tostring())
						$runspaces.Add($temp) | Out-Null
						
						#loop through existing runspaces one time
						Get-RunspaceData
						
						#If we have more running than max queue (used to control timeout accuracy)
						#Script scope resolves odd PowerShell 2 issue
						$firstRun = $true
						while ($runspaces.count -ge $Script:MaxQueue)
						{
							
							#give verbose output
							if ($firstRun)
							{
								Write-Verbose "$($runspaces.count) items running - exceeded $Script:MaxQueue limit."
							}
							$firstRun = $false
							
							#run get-runspace data and sleep for a short while
							Get-RunspaceData
							Start-Sleep -Milliseconds $sleepTimer
							
						}
						
						#endregion add scripts to runspace pool
					}
					
					Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f (@($runspaces | Where { $_.Runspace -ne $Null }).Count))
					Get-RunspaceData -wait
					
					if (-not $quiet)
					{
						Write-Progress -Activity "Running Query" -Status "Starting threads" -Completed
					}
					
				}
				Finally
				{
					#Close the runspace pool, unless we specified no close on timeout and something timed out
					if (($timedOutTasks -eq $false) -or (($timedOutTasks -eq $true) -and ($noCloseOnTimeout -eq $false)))
					{
						Write-Verbose "Closing the runspace pool"
						$runspacepool.close()
					}
					
					#collect garbage
					[gc]::Collect()
				}
			}
		}
		
		Write-Verbose "PSBoundParameters = $($PSBoundParameters | Out-String)"
		
		$bound = $PSBoundParameters.keys -contains "ComputerName"
		if (-not $bound)
		{
			[System.Collections.ArrayList]$AllComputers = @()
		}
	}
	Process
	{
		
		#Handle both pipeline and bound parameter.  We don't want to stream objects, defeats purpose of parallelizing work
		if ($bound)
		{
			$AllComputers = $ComputerName
		}
		Else
		{
			foreach ($Computer in $ComputerName)
			{
				$AllComputers.add($Computer) | Out-Null
			}
		}
		
	}
	End
	{
		
		#Built up the parameters and run everything in parallel
		$params = @($Detail, $Quiet)
		$splat = @{
			Throttle = $Throttle
			RunspaceTimeout = $Timeout
			InputObject = $AllComputers
			parameter = $params
		}
		if ($NoCloseOnTimeout)
		{
			$splat.add('NoCloseOnTimeout', $True)
		}
		
		Invoke-Parallel @splat -ScriptBlock {
			
			$computer = $_.trim()
			$detail = $parameter[0]
			$quiet = $parameter[1]
			
			#They want detail, define and run test-server
			if ($detail)
			{
				Try
				{
					#Modification of jrich's Test-Server Function: https://gallery.technet.microsoft.com/scriptcenter/Powershell-Test-Server-e0cdea9a
					Function Test-Server
					{
						[cmdletBinding()]
						param (
							[parameter(
									   Mandatory = $true,
									   ValueFromPipeline = $true)]
							[string[]]$ComputerName,
							
							[switch]$All,
							
							[parameter(Mandatory = $false)]
							[switch]$CredSSP,
							
							[switch]$RemoteReg,
							
							[switch]$RDP,
							
							[switch]$RPC,
							
							[switch]$SMB,
							
							[switch]$WSMAN,
							
							[switch]$IPV6,
							
							[Management.Automation.PSCredential]$Credential
						)
						begin
						{
							$total = Get-Date
							$results = @()
							if ($credssp -and -not $Credential)
							{
								Throw "Must supply Credentials with CredSSP test"
							}
							
							[string[]]$props = write-output Name, IP, Domain, Ping, WSMAN, CredSSP, RemoteReg, RPC, RDP, SMB
							
							#Hash table to create PSObjects later, compatible with ps2...
							$Hash = @{ }
							foreach ($prop in $props)
							{
								$Hash.Add($prop, $null)
							}
							
							Function Test-Port
							{
								[cmdletbinding()]
								Param (
									[string]$srv,
									
									$port = 135,
									
									$timeout = 3000
								)
								$ErrorActionPreference = "SilentlyContinue"
								$tcpclient = new-Object system.Net.Sockets.TcpClient
								$iar = $tcpclient.BeginConnect($srv, $port, $null, $null)
								$wait = $iar.AsyncWaitHandle.WaitOne($timeout, $false)
								if (-not $wait)
								{
									$tcpclient.Close()
									Write-Verbose "Connection Timeout to $srv`:$port"
									$false
								}
								else
								{
									Try
									{
										$tcpclient.EndConnect($iar) | out-Null
										$true
									}
									Catch
									{
										write-verbose "Error for $srv`:$port`: $_"
										$false
									}
									$tcpclient.Close()
								}
							}
						}
						
						process
						{
							foreach ($name in $computername)
							{
								$dt = $cdt = Get-Date
								Write-verbose "Testing: $Name"
								$failed = 0
								try
								{
									$DNSEntity = [Net.Dns]::GetHostEntry($name)
									$domain = ($DNSEntity.hostname).replace("$name.", "")
									$ips = $DNSEntity.AddressList | %{
										if (-not (-not $IPV6 -and $_.AddressFamily -like "InterNetworkV6"))
										{
											$_.IPAddressToString
										}
									}
								}
								catch
								{
									$rst = New-Object -TypeName PSObject -Property $Hash | Select -Property $props
									$rst.name = $name
									$results += $rst
									$failed = 1
								}
								Write-verbose "DNS:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
								if ($failed -eq 0)
								{
									foreach ($ip in $ips)
									{
										
										$rst = New-Object -TypeName PSObject -Property $Hash | Select -Property $props
										$rst.name = $name
										$rst.ip = $ip
										$rst.domain = $domain
										
										if ($RDP -or $All)
										{
											####RDP Check (firewall may block rest so do before ping
											try
											{
												$socket = New-Object Net.Sockets.TcpClient($name, 3389) -ErrorAction stop
												if ($socket -eq $null)
												{
													$rst.RDP = $false
												}
												else
												{
													$rst.RDP = $true
													$socket.close()
												}
											}
											catch
											{
												$rst.RDP = $false
												Write-Verbose "Error testing RDP: $_"
											}
										}
										Write-verbose "RDP:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
										#########ping
										if (test-connection $ip -count 2 -Quiet)
										{
											Write-verbose "PING:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
											$rst.ping = $true
											
											if ($WSMAN -or $All)
											{
												try
												{
													############wsman
														Test-WSMan $ip -ErrorAction stop | Out-Null
														$rst.WSMAN = $true
													}
													catch
													{
														$rst.WSMAN = $false
														Write-Verbose "Error testing WSMAN: $_"
													}
													Write-verbose "WSMAN:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													if ($rst.WSMAN -and $credssp) ########### credssp
													{
														try
														{
															Test-WSMan $ip -Authentication Credssp -Credential $cred -ErrorAction stop
															$rst.CredSSP = $true
														}
														catch
														{
															$rst.CredSSP = $false
															Write-Verbose "Error testing CredSSP: $_"
														}
														Write-verbose "CredSSP:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													}
												}
												if ($RemoteReg -or $All)
												{
													try ########remote reg
													{
														[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ip) | Out-Null
														$rst.remotereg = $true
													}
													catch
													{
														$rst.remotereg = $false
														Write-Verbose "Error testing RemoteRegistry: $_"
													}
													Write-verbose "remote reg:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
												}
												if ($RPC -or $All)
												{
													try ######### wmi
													{
														$w = [wmi] ''
														$w.psbase.options.timeout = 15000000
														$w.path = "\\$Name\root\cimv2:Win32_ComputerSystem.Name='$Name'"
														$w | select none | Out-Null
														$rst.RPC = $true
													}
													catch
													{
														$rst.rpc = $false
														Write-Verbose "Error testing WMI/RPC: $_"
													}
													Write-verbose "WMI/RPC:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
												}
												if ($SMB -or $All)
												{
													
													#Use set location and resulting errors.  push and pop current location
													try ######### C$
													{
														$path = "\\$name\c$"
														Push-Location -Path $path -ErrorAction stop
														$rst.SMB = $true
														Pop-Location
													}
													catch
													{
														$rst.SMB = $false
														Write-Verbose "Error testing SMB: $_"
													}
													Write-verbose "SMB:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													
												}
											}
											else
											{
												$rst.ping = $false
												$rst.wsman = $false
												$rst.credssp = $false
												$rst.remotereg = $false
												$rst.rpc = $false
												$rst.smb = $false
											}
											$results += $rst
										}
									}
									Write-Verbose "Time for $($Name): $((New-TimeSpan $cdt ($dt)).totalseconds)"
									Write-Verbose "----------------------------"
								}
							}
							end
							{
								Write-Verbose "Time for all: $((New-TimeSpan $total ($dt)).totalseconds)"
								Write-Verbose "----------------------------"
								return $results
							}
						}
						
						#Build up parameters for Test-Server and run it
						$TestServerParams = @{
							ComputerName = $Computer
							ErrorAction = "Stop"
						}
						
						if ($detail -eq "*")
						{
							$detail = "WSMan", "RemoteReg", "RPC", "RDP", "SMB"
						}
						
						$detail | Select -Unique | Foreach-Object { $TestServerParams.add($_, $True) }
						Test-Server @TestServerParams | Select -Property $("Name", "IP", "Domain", "Ping" + $detail)
					}
					Catch
					{
						Write-Warning "Error with Test-Server: $_"
					}
				}
				#We just want ping output
				else
				{
					Try
					{
						#Pick out a few properties, add a status label.  If quiet output, just return the address
						$result = $null
						if ($result = @(Test-Connection -ComputerName $computer -Count 2 -erroraction Stop))
						{
							$Output = $result | Select -first 1 -Property Address,
													   IPV4Address,
													   IPV6Address,
													   ResponseTime,
													   @{ label = "STATUS"; expression = { "Responding" } }
							
							if ($quiet)
							{
								$Output.address
							}
							else
							{
								$Output
							}
						}
					}
					Catch
					{
						if (-not $quiet)
						{
							#Ping failed.  I'm likely making inappropriate assumptions here, let me know if this is the case : )
							if ($_ -match "No such host is known")
							{
								$status = "Unknown host"
							}
							elseif ($_ -match "Error due to lack of resources")
							{
								$status = "No Response"
							}
							else
							{
								$status = "Error: $_"
							}
							
							"" | Select -Property @{ label = "Address"; expression = { $computer } },
										IPV4Address,
										IPV6Address,
										ResponseTime,
										@{ label = "STATUS"; expression = { $status } }
						}
					}
				}
			}
		}
	}
# End Invoke-Ping
## Begin Set-DnsServerIpAddress
Function Set-DnsServerIpAddress {
    param(
        [string] $ComputerName,
        [string] $NicName,
        [string] $IpAddresses
    )
    if (Test-Connection -ComputerName $ComputerName -Count 2 -Quiet) {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock { param ($ComputerName, $NicName, $IpAddresses)
            write-host "Setting on $ComputerName on interface $NicName a new set of DNS Servers $IpAddresses"
            Set-DnsClientServerAddress -InterfaceAlias $NicName -ServerAddresses $IpAddresses
        } -ArgumentList $ComputerName, $NicName, $IpAddresses
    } else {
        write-host "Can't access $ComputerName. Computer is not online."
    }
}
## End Set-DnsServerIpAddress
## Begin WhoIs
Function WhoIs {
param (
                [Parameter(Mandatory=$True,
                           HelpMessage='Please enter domain name (e.g. microsoft.com)')]
                           [string]$domain
        )
Write-Host "Connecting to Web Services URL..." -ForegroundColor Green
try {
#Retrieve the data from web service WSDL
If ($whois = New-WebServiceProxy -uri "http://www.webservicex.net/whois.asmx?WSDL") {Write-Host "Ok" -ForegroundColor Green}
else {Write-Host "Error" -ForegroundColor Red}
Write-Host "Gathering $domain data..." -ForegroundColor Green
#Return the data
(($whois.getwhois("=$domain")).Split("<<<")[0])
} catch {

Write-Host "Please enter valid domain name (e.g. microsoft.com)." -ForegroundColor Red}
}
## End WhoIs
## Begin Get-NetworkStatistics
Function Get-NetworkStatistics{ 
    $properties = 'Protocol','LocalAddress','LocalPort' 
    $properties += 'RemoteAddress','RemotePort','State','ProcessName','PID' 

    netstat -ano | Select-String -Pattern '\s+(TCP|UDP)' | ForEach-Object { 

        $item = $_.line.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries) 

        if($item[1] -notmatch '^\[::') 
        {            
            if (($la = $item[1] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6') 
            { 
               $localAddress = $la.IPAddressToString 
               $localPort = $item[1].split('\]:')[-1] 
            } 
            else 
            { 
                $localAddress = $item[1].split(':')[0] 
                $localPort = $item[1].split(':')[-1] 
            }  

            if (($ra = $item[2] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6') 
            { 
               $remoteAddress = $ra.IPAddressToString 
               $remotePort = $item[2].split('\]:')[-1] 
            } 
            else 
            { 
               $remoteAddress = $item[2].split(':')[0] 
               $remotePort = $item[2].split(':')[-1] 
            }  

            New-Object PSObject -Property @{ 
                PID = $item[-1] 
                ProcessName = (Get-Process -Id $item[-1] -ErrorAction SilentlyContinue).Name 
                Protocol = $item[0] 
                LocalAddress = $localAddress 
                LocalPort = $localPort 
                RemoteAddress =$remoteAddress 
                RemotePort = $remotePort 
                State = if($item[0] -eq 'tcp') {$item[3]} else {$null} 
            } | Select-Object -Property $properties 
        } 
    } 
}
## End Get-NetworkStatistics
## Begin Test-Port
Function Test-Port{    
[cmdletbinding(    
    DefaultParameterSetName = '',    
    ConfirmImpact = 'low'    
)]    
    Param(    
        [Parameter(    
            Mandatory = $True,    
            Position = 0,    
            ParameterSetName = '',    
            ValueFromPipeline = $True)]    
            [array]$computer,    
        [Parameter(    
            Position = 1,    
            Mandatory = $True,    
            ParameterSetName = '')]    
            [array]$port,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [int]$TCPtimeout=1000,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [int]$UDPtimeout=1000,               
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [switch]$TCP,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [switch]$UDP                                      
        )    
    Begin {    
        If (!$tcp -AND !$udp) {$tcp = $True}    
        #Typically you never do this, but in this case I felt it was for the benefit of the Function    
        #as any errors will be noted in the output of the report            
        $ErrorActionPreference = "SilentlyContinue"    
        $report = @()    
    }    
    Process {       
        ForEach ($c in $computer) {    
            ForEach ($p in $port) {    
                If ($tcp) {      
                    #Create temporary holder     
                    $temp = "" | Select Server, Port, TypePort, Open, Notes    
                    #Create object for connecting to port on computer    
                    $tcpobject = new-Object system.Net.Sockets.TcpClient    
                    #Connect to remote machine's port                  
                    $connect = $tcpobject.BeginConnect($c,$p,$null,$null)    
                    #Configure a timeout before quitting    
                    $wait = $connect.AsyncWaitHandle.WaitOne($TCPtimeout,$false)    
                    #If timeout    
                    If(!$wait) {    
                        #Close connection    
                        $tcpobject.Close()    
                        Write-Verbose "Connection Timeout"    
                        #Build report    
                        $temp.Server = $c    
                        $temp.Port = $p    
                        $temp.TypePort = "TCP"    
                        $temp.Open = "False"    
                        $temp.Notes = "Connection to Port Timed Out"    
                    } Else {    
                        $error.Clear()    
                        $tcpobject.EndConnect($connect) | out-Null    
                        #If error    
                        If($error[0]){    
                            #Begin making error more readable in report    
                            [string]$string = ($error[0].exception).message    
                            $message = (($string.split(":")[1]).replace('"',"")).TrimStart()    
                            $failed = $true    
                        }    
                        #Close connection        
                        $tcpobject.Close()    
                        #If unable to query port to due failure    
                        If($failed){    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "TCP"    
                            $temp.Open = "False"    
                            $temp.Notes = "$message"    
                        } Else{    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "TCP"    
                            $temp.Open = "True"      
                            $temp.Notes = ""    
                        }    
                    }       
                    #Reset failed value    
                    $failed = $Null        
                    #Merge temp array with report                
                    $report += $temp    
                }        
                If ($udp) {    
                    #Create temporary holder     
                    $temp = "" | Select Server, Port, TypePort, Open, Notes                                       
                    #Create object for connecting to port on computer    
                    $udpobject = new-Object system.Net.Sockets.Udpclient  
                    #Set a timeout on receiving message   
                    $udpobject.client.ReceiveTimeout = $UDPTimeout   
                    #Connect to remote machine's port                  
                    Write-Verbose "Making UDP connection to remote server"   
                    $udpobject.Connect("$c",$p)   
                    #Sends a message to the host to which you have connected.   
                    Write-Verbose "Sending message to remote host"   
                    $a = new-object system.text.asciiencoding   
                    $byte = $a.GetBytes("$(Get-Date)")   
                    [void]$udpobject.Send($byte,$byte.length)   
                    #IPEndPoint object will allow us to read datagrams sent from any source.    
                    Write-Verbose "Creating remote endpoint"   
                    $remoteendpoint = New-Object system.net.ipendpoint([system.net.ipaddress]::Any,0)   
                    Try {   
                        #Blocks until a message returns on this socket from a remote host.   
                        Write-Verbose "Waiting for message return"   
                        $receivebytes = $udpobject.Receive([ref]$remoteendpoint)   
                        [string]$returndata = $a.GetString($receivebytes)  
                        If ($returndata) {  
                           Write-Verbose "Connection Successful"    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "UDP"    
                            $temp.Open = "True"    
                            $temp.Notes = $returndata     
                            $udpobject.close()     
                        }                         
                    } Catch {   
                        If ($Error[0].ToString() -match "\bRespond after a period of time\b") {   
                            #Close connection    
                            $udpobject.Close()    
                            #Make sure that the host is online and not a false positive that it is open   
                            If (Test-Connection -comp $c -count 1 -quiet) {   
                                Write-Verbose "Connection Open"    
                                #Build report    
                                $temp.Server = $c    
                                $temp.Port = $p    
                                $temp.TypePort = "UDP"    
                                $temp.Open = "True"    
                                $temp.Notes = ""   
                            } Else {   
                                <#   
                                It is possible that the host is not online or that the host is online,    
                                but ICMP is blocked by a firewall and this port is actually open.   
                                #>   
                                Write-Verbose "Host maybe unavailable"    
                                #Build report    
                                $temp.Server = $c    
                                $temp.Port = $p    
                                $temp.TypePort = "UDP"    
                                $temp.Open = "False"    
                                $temp.Notes = "Unable to verify if port is open or if host is unavailable."                                   
                            }                           
                        } ElseIf ($Error[0].ToString() -match "forcibly closed by the remote host" ) {   
                            #Close connection    
                            $udpobject.Close()    
                            Write-Verbose "Connection Timeout"    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "UDP"    
                            $temp.Open = "False"    
                            $temp.Notes = "Connection to Port Timed Out"                           
                        } Else {                        
                            $udpobject.close()   
                        }   
                    }       
                    #Merge temp array with report                
                    $report += $temp    
                }                                    
            }    
        }                    
    }    
    End {    
        #Generate Report    
        $report   
    }  
} 
## End Test-Port
## Begin Get-IPAddress
Function Get-IPAddress{
	Get-NetIPAddress | ?{($_.interfacealias -notlike "*loopback*")`
 -and ($_.interfacealias -notlike "*vmware*")`
  -and ($_.interfacealias -notlike "*loopback*")`
   -and ($_.interfacealias -notlike "*bluetooth*")`
    -and ($_.interfacealias -notlike "*isatap*")} | ft
}
## End Get-IPAddress
## Begin Get-IPv6InWindows
Function Get-IPv6InWindows{
   <#
         .SYNOPSIS
         Get the configured IPv6 value from the registry

         .DESCRIPTION
         Get the configured IPv6 value from the registry
         Transforms the Registry value into human understandable values

         .EXAMPLE
         PS C:\> Get-IPv6InWindows
         All IPv6 components are enabled (0)

         .EXAMPLE
         PS C:\> Get-IPv6InWindows -verbose
         Prefer IPv4 over IPv6 (32)

         Get the configured IPv6 value from the registry, with verbose output

         .LINK
         Set-IPv6InWindows

         .LINK
         https://docs.microsoft.com/en-us/troubleshoot/windows-server/networking/configure-ipv6-in-windows

         .LINK
         https://docs.microsoft.com/en-us/troubleshoot/windows-server/networking/configure-ipv6-in-windows#reference

         .NOTES
         Just a wrapper to make the values more human readable.
         This is just a quick and dirty initial version!

         If you find any further values (other then the supported), please let me know!

         Want to modify your IPv6 configuration? Use its companion Set-IPv6InWindows
   #>
   [CmdletBinding(ConfirmImpact = 'None')]
   [OutputType([string])]
   param ()

   begin
   {
      # Cleanup
      $ComponentValue = $null
      $ComponentValueText = $null

      #region BoundParameters
      if (($PSCmdlet.MyInvocation.BoundParameters['Verbose']).IsPresent)
      {
         $IsVerbose = $true
      }
      else
      {
         $IsVerbose = $false
      }

      if (($PSCmdlet.MyInvocation.BoundParameters['Debug']).IsPresent)
      {
         $IsDebug = $true
      }
      else
      {
         $IsDebug = $false
      }
      #endregion BoundParameters
   }

   process
   {
      # Get the Value from the registry
      try
      {
         $paramGetItemProperty = @{
            Path          = 'HKLM:\SYSTEM\CurrentControlSet\Services\tcpip6\Parameters'
            Name          = 'DisabledComponents'
            Debug         = $IsDebug
            Verbose       = $IsVerbose
            ErrorAction   = 'Stop'
            WarningAction = 'Continue'
         }
         $ComponentValue = (Get-ItemProperty @paramGetItemProperty | Select-Object -ExpandProperty DisabledComponents -ErrorAction Stop -WarningAction Continue)
      }
      catch
      {
         #region ErrorHandler
         # get error record
         [Management.Automation.ErrorRecord]$e = $_

         # retrieve information about runtime error
         $info = [PSCustomObject]@{
            Exception = $e.Exception.Message
            Reason    = $e.CategoryInfo.Reason
            Target    = $e.CategoryInfo.TargetName
            Script    = $e.InvocationInfo.ScriptName
            Line      = $e.InvocationInfo.ScriptLineNumber
            Column    = $e.InvocationInfo.OffsetInLine
         }

         Write-Verbose -Message $info

         Write-Error -Message ($info.Exception) -ErrorAction Stop

         # Only here to catch a global ErrorAction overwrite
         exit 1
         #endregion ErrorHandler
      }

      switch ($ComponentValue)
      {
         0
         {
            $ComponentValueText = ('All IPv6 components are enabled ({0})' -f $ComponentValue)
         }
         255
         {
            $ComponentValueText = ('All IPv6 components are disabled ({0})' -f $ComponentValue)
         }
         2
         {
            $ComponentValueText = ('6to4 is disabled ({0})' -f $ComponentValue)
         }
         4
         {
            $ComponentValueText = ('ISATAP is disabled ({0})' -f $ComponentValue)
         }
         8
         {
            $ComponentValueText = ('Teredo is disabled ({0})' -f $ComponentValue)
         }
         10
         {
            $ComponentValueText = ('Teredo and 6to4 is disabled ({0})' -f $ComponentValue)
         }
         1
         {
            $ComponentValueText = ('All tunnel interfaces are disabled ({0})' -f $ComponentValue)
         }
         16
         {
            $ComponentValueText = ('All LAN and PPP interfaces are disabled ({0})' -f $ComponentValue)
         }
         17
         {
            $ComponentValueText = ('All LAN, PPP and tunnel interfaces are disabled ({0})' -f $ComponentValue)
         }
         32
         {
            $ComponentValueText = ('Prefer IPv4 over IPv6 ({0})' -f $ComponentValue)
         }
         default
         {
            $ComponentValueText = ('Unknown value found: {0}' -f $ComponentValue)
         }
      }
   }

   end
   {
      # Dump the info
      $ComponentValueText
   }
}
## End Get-IPv6InWindows
## Begin Get-IPAddress
Function Get-IPAddress{
	Get-NetIPAddress | ?{($_.interfacealias -notlike "*loopback*") -and ($_.interfacealias -notlike "*vmware*") -and ($_.interfacealias -notlike "*loopback*") -and ($_.interfacealias -notlike "*bluetooth*") -and ($_.interfacealias -notlike "*isatap*")} | ft
}
## End Get-IPAddress
## Begin Get-SnmpTrap
Function Get-SnmpTrap {
<#
.SYNOPSIS
## Begin that
Function that will list SNMP Community string, Security Options and Trap Configuration for SNMP version 1 and version 2c.
.DESCRIPTION
** This Function will list SNMP settings of  windows server by reading the registry keys under HKLM\SYSTEM\CurrentControlSet\services\SNMP\Parameters\
Example usage:																					  
Get_SnmpTrap
This will list the  SNMP Community string, Security Options and Trap Configuration on the server. The meaning of each column is:
AcceptedCommunityStrings => The community string that the SNMP agent is allowed to receive. If the host is not requested with one of these pre-defined 
community strings, then the host will send an authentication trap.
AllowedHosts => The hostnames or IP addresses from which SNMP agent will accept SNMP messages.
CommunityRights => The permission that determines how the SNMP agent processes the incoming request from various communities.
TrapCommunityNames => When an SNMP agent receives a request that does not contain a valid community name or the host that is sending the message 
is not on the list of acceptable hosts, the agent can send an authentication trap message to one or more trap destinations (management systems)
TrapDestinations => The host names or IP addresses of trap destinations which are defined under the TrapCommunityNames.
SendTrap => It indicates whether sending autentication trap is enabled.
Author: phyoepaing3.142@gmail.com
Country: Myanmar(Burma)
Released: 05/07/2017
.EXAMPLE
Get_SnmpTrap
This will list the  SNMP Community string, Security Options and Trap Configuration on the server.
.LINK
You can find this script and more at: https://www.sysadminplus.blogspot.com/
#>

### DATA lookup section to convert registry numeric to corresponding output ###
$ConvertRights = DATA { ConvertFrom-StringData -StringData @'
1 = NONE
2 = NOTIFY
4 = READ-ONLY
8 = READ-WRITE
16 = READ-CREATE
'@}

$rh = '2147483650';  ## This number represents HKLM
$key1 = 'SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers';
$reg = [wmiclass]"\\localhost\root\default:StdRegprov"; 
$obj = New-Object -TypeName PsObject -Property @{AllowedHosts=@(); AcceptedCommunityStrings="";  CommunityRights =@(); TrapCommunityNames=@(); TrapDestinations=@(); SendTrap="" }; 
$AccessDenied = 0;

### Read the registry to find the allowed hosts for incoming community string ###
$i=1;
while ( $reg.GetStringValue($rh, $key1, $i ).sValue )
	{
	$obj.AllowedHosts += $reg.GetStringValue($rh, $key1, $i ).sValue; 
	$i ++;
	}
If ($obj.AllowedHosts.count -eq 1)
	{
	$obj.AllowedHosts = $obj.AllowedHosts[0];
	}

### Read the Community Strings ###	
Try {
	$obj.AcceptedCommunityStrings = (Gi -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities -EA Stop).Property; 

## If there is only one community string, then convert the property type to string from array ###
	If ($obj.AcceptedCommunityStrings.count -eq 1)
	{ $obj.AcceptedCommunityStrings = $obj.AcceptedCommunityStrings[0] }

### If there are multiple community strings, then read through all the security permission of each community string 	via registry ##
	If ($obj.AcceptedCommunityStrings -is [array])
	{	
		$obj.AcceptedCommunityStrings | foreach {
		$securityRight =  [string]((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities).$_)
		$obj.CommunityRights += $_+":"+$ConvertRights[$securityRight]
			}
		}
	else
		{
		[string]$securityRight = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities).$($obj.AcceptedCommunityStrings)
		$obj.CommunityRights = $ConvertRights[$securityRight]
		}
	}
catch [System.Security.SecurityException]
	{ 
	Write-Host -fore red "Access to Registry is denied. Please make sure you have permission to access registry and run in the elevated command prompt.`n"; 
	$obj.AllowedHosts = "N/A"
	$obj.AcceptedCommunityStrings = "N/A"
	$obj.CommunityRights = "N/A"
	$obj.TrapCommunityNames = "N/A"
	$obj.TrapDestinations = "N/A"
	$obj.SendTrap = "N/A"
	$AccessDenied = 1; 
	}
catch 
	{ 
	Write-Host -fore red "SMNP Service is not installed on one or more servers.`n"; 
	$obj.AllowedHosts = "N/A"
	$obj.AcceptedCommunityStrings = "N/A"
	$obj.CommunityRights = "N/A"
	$obj.TrapCommunityNames = "N/A"
	$obj.TrapDestinations = "N/A"
	$obj.SendTrap = "N/A"
	$AccessDenied = 1; 
	$obj;
	}

## If the read of registry is not access-denied from previous try-catch statement, then continue ##	
If (!$AccessDenied)	
	{
	Try {
		$TrapConfig = Gci -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration -EA Stop ;
		$TrapConfig | foreach {
			$obj.TrapCommunityNames += $_.PsChildName
			}
	If ($obj.TrapCommunityNames.count -eq 1)	
	{ $obj.TrapCommunityNames = $obj.TrapCommunityNames[0]	}
			
		}
	catch 
		{  }
		
### Find destination for each Trap. The trap's community name will be prefixed on the trap's destination IP/hosts if there are multiple Traps configured, if it's single trap, then use without prefix ###

If ($obj.TrapCommunityNames -is [array])
	{
		$obj.TrapCommunityNames | foreach {
		$key2 = "SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\$_";
		$i=1;
			while ( $reg.GetStringValue($rh, $key2, $i ).sValue )
				{
				$obj.TrapDestinations += $_+":"+$reg.GetStringValue($rh, $key2, $i ).sValue; 
				$i ++;
				}
		}
	}
else
	{
	$key2 = "SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\$($obj.TrapCommunityNames)";
	$i=1;
		while ( $reg.GetStringValue($rh, $key2, $i ).sValue )
				{
				$obj.TrapDestinations += $reg.GetStringValue($rh, $key2, $i ).sValue; 
				$i ++;
				}
		}
	
### If there is only one entry in the Trap Destination, then convert  the array to string ###
If ($obj.TrapDestinations.count -eq 1)
	{
	$obj.TrapDestinations = $obj.TrapDestinations[0];
	}
	
#### Check if the 'Send Authentication Trap' check box is enabled ###
	Switch ((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters).EnableAuthenticationTraps)
		{
		"0" { $obj.SendTrap = "Disabled" }
		"1" { $obj.SendTrap = "Enabled "}		
		}
	$obj;	
	}
}
## End Get-SnmpTrap