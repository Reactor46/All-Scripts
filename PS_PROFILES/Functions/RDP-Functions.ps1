## Begin Set-RDPDisable
Function Set-RDPDisable{
<#
	.SYNOPSIS
		The Function Set-RDPDisable disable RDP remotely using the registry
	
	.DESCRIPTION
		The Function Set-RDPDisable disable RDP remotely using the registry
	
	.PARAMETER ComputerName
		Specifies the ComputerName
	
	.EXAMPLE
		PS C:\> Set-RDPDisable
	
	.EXAMPLE
		PS C:\> Set-RDPDisable -ComputerName "DC01"
	
	.EXAMPLE
		PS C:\> Set-RDPDisable -ComputerName "DC01","DC02","DC03"
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	[CmdletBinding()]
	PARAM (
		[String[]]$ComputerName = $env:COMPUTERNAME
	)
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				IF (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
				{
					$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer)
					$regKey = $regKey.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Terminal Server", $True)
					$regkey.SetValue("fDenyTSConnections", 1)
					$regKey.flush()
					$regKey.Close()
				} #IF Test-Connection
			} #Try
			CATCH
			{
				$Error[0].Exception.Message
			} #Catch
		} #FOREACH
	} #Process
}
## End Set-RDPDisable
## Begin Set-RDPEnable
Function Set-RDPEnable{
<#
	.SYNOPSIS
		The Function Set-RDPEnable enable RDP remotely using the registry
	
	.DESCRIPTION
		The Function Set-RDPEnable enable RDP remotely using the registry
	
	.PARAMETER ComputerName
		Specifies the ComputerName
	
	.EXAMPLE
		PS C:\> Set-RDPEnable
	
	.EXAMPLE
		PS C:\> Set-RDPEnable -ComputerName "DC01"
	
	.EXAMPLE
		PS C:\> Set-RDPEnable -ComputerName "DC01","DC02","DC03"
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	
	[CmdletBinding()]
	PARAM (
		[String[]]$ComputerName = $env:COMPUTERNAME
	)
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				IF (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
				{
					$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer)
					$regKey = $regKey.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Terminal Server", $True)
					$regkey.SetValue("fDenyTSConnections", 0)
					$regKey.flush()
					$regKey.Close()
				} #IF Test-Connection
			} #Try
			CATCH
			{
				$Error[0].Exception.Message
			} #Catch
		} #FOREACH
	} #Process
}
## End Set-RDPEnable
## Begin Set-RemoteDesktop
Function Set-RemoteDesktop{
<#
	.SYNOPSIS
		The Function Set-RemoteDesktop allows you to enable or disable RDP remotely using the registry
	
	.DESCRIPTION
		The Function Set-RemoteDesktop allows you to enable or disable RDP remotely using the registry
	
	.PARAMETER ComputerName
		Specifies the ComputerName
	
	.EXAMPLE
		PS C:\> Set-RemoteDesktop -enable $true
	
	.EXAMPLE
		PS C:\> Set-RemoteDesktop -ComputerName "DC01" -enable $false
	
	.EXAMPLE
		PS C:\> Set-RemoteDesktop -ComputerName "DC01","DC02","DC03" -enable $false
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	
	[CmdletBinding()]
	PARAM (
		[String[]]$ComputerName = $env:COMPUTERNAME,
		[Parameter(Mandatory = $true)]	
		[Boolean]$Enable
	)
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				IF (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
				{
					$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer)
					$regKey = $regKey.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Terminal Server", $True)
					
					IF ($Enable){$regkey.SetValue("fDenyTSConnections", 0)}
					ELSE { $regkey.SetValue("fDenyTSConnections", 1)}
					$regKey.flush()
					$regKey.Close()
				} #IF Test-Connection
			} #Try
			CATCH
			{
				$Error[0].Exception.Message
			} #Catch
		} #FOREACH
	} #Process
}
## End Set-RemoteDesktop
## Begin Start-RDP
Function Start-RDP {

    [CmdletBinding()]
    param(
    [parameter(
        mandatory,
        position = 0)]
    [string]
    $server,

    [parameter(
        mandatory = 0,
        position = 1)]
    [string]
    $RDPPort = 3389
    )

    # Test RemoteDesktop Connection is valid or not
    $TestRemoteDesktop = New-Object System.Net.Sockets.TCPClient -ArgumentList $server,$RDPPort

    # Execute RDP Connection
    if ($TestRemoteDesktop)
    {
        Invoke-Expression "mstsc /v:$server"
    }
    else
    {
        Write-Warning "RemoteDesktop"
    }

}
## End Start-RDP
## Begin RDP
Function RDP {
  <# 
  .SYNOPSIS 
  Remote Desktop Protocol to specified workstation(s) 

  .EXAMPLE 
  RDP Computer123456 

  .EXAMPLE 
  RDP 123456 
  #> 
	param(
	[Parameter(Mandatory=$true)]
	[string]$computername)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

	#Start Remote Desktop Protocol on specifed workstation
	& "C:\windows\system32\mstsc.exe" /v:$computername /fullscreen
}
## End RDP
## Begin Disable-RemoteDesktop
Function Disable-RemoteDesktop{
<#
	.SYNOPSIS
		The Function Disable-RemoteDesktop will disable RemoteDesktop on a local or remote machine.
	
	.DESCRIPTION
		The Function Disable-RemoteDesktop will disable RemoteDesktop on a local or remote machine.
	
	.PARAMETER ComputerName
		Specifies the computername
	
	.PARAMETER Credential
		Specifies the credential to use
	
	.PARAMETER CimSession
		Specifies one or more existing CIM Session(s) to use
	
	.EXAMPLE
		PS C:\> Disable-RemoteDesktop -ComputerName DC01
	
	.EXAMPLE
		PS C:\> Disable-RemoteDesktop -ComputerName DC01 -Credential (Get-Credential -cred "FX\SuperAdmin")
	
	.EXAMPLE
		PS C:\> Disable-RemoteDesktop -CimSession $Session
	
	.EXAMPLE
		PS C:\> Disable-RemoteDesktop -CimSession $Session1,$session2,$session3
	
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		www.lazywinadmin.com
        github.com/lazywinadmin
#>
	#Requires -RunAsAdministrator
	[CmdletBinding(DefaultParameterSetName = 'CimSession',
				   SupportsShouldProcess = $true)]
	PARAM (
		[Parameter(
				   ParameterSetName = "Main",
				   ValueFromPipeline = $True,
				   ValueFromPipelineByPropertyName = $True)]
		[Alias("CN", "__SERVER", "PSComputerName")]
		[String[]]$ComputerName,
		
		[Parameter(ParameterSetName = "Main")]
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty,
		
		[Parameter(ParameterSetName = "CimSession")]
		[Microsoft.Management.Infrastructure.CimSession[]]$CimSession
	)
	BEGIN
	{
		# Helper Function

		Function Get-DefaultMessage
		{
<#
.SYNOPSIS
	Helper Function to show default message used in VERBOSE/DEBUG/WARNING
.DESCRIPTION
	Helper Function to show default message used in VERBOSE/DEBUG/WARNING
	and... HOST in some case.
	This is helpful to standardize the output messages
	
.PARAMETER Message
	Specifies the message to show
.NOTES
	Francois-Xavier Cat
	www.lazywinadmin.com
	@lazywinadm
#>
			PARAM ($Message)
			$DateFormat = Get-Date -Format 'yyyy/MM/dd-HH:mm:ss:ff'
			$FunctionName = (Get-Variable -Scope 1 -Name MyInvocation -ValueOnly).MyCommand.Name
			Write-Output "[$DateFormat][$FunctionName] $Message"
		} #Get-DefaultMessage
	}
	PROCESS
	{
		IF ($PSBoundParameters['CimSession'])
		{
			FOREACH ($Cim in $CimSession)
			{
				$CIMComputer = $($Cim.ComputerName).ToUpper()
				
				IF ($PSCmdlet.ShouldProcess($CIMComputer, "Disable Remote Desktop via Win32_TerminalServiceSetting"))
				{
					
					TRY
					{
						# Parameters for Get-CimInstance
						$CIMSplatting = @{
							Class = "Win32_TerminalServiceSetting"
							NameSpace = "root\cimv2\terminalservices"
							CimSession = $Cim
							ErrorAction = 'Stop'
							ErrorVariable = "ErrorProcessGetCimInstance"
						}
						
						# Parameters for Invoke-CimMethod
						$CIMInvokeSplatting = @{
							MethodName = "SetAllowTSConnections"
							Arguments = @{
								AllowTSConnections = 0;
								ModifyFirewallException = 0
							}
							ErrorAction = 'Stop'
							ErrorVariable = "ErrorProcessInvokeCim"
						}
						
						Write-Verbose -Message (Get-DefaultMessage -Message "$CIMComputer - CIMSession - disable Remote Desktop (and Modify Firewall Exception")
						Get-CimInstance @CIMSplatting | Invoke-CimMethod @CIMInvokeSplatting
					}
					CATCH
					{
						Write-Warning -Message (Get-DefaultMessage -Message "$CIMComputer - CIMSession - Something wrong happened")
						IF ($ErrorProcessGetCimInstance) { Write-Warning -Message (Get-DefaultMessage -Message "$CIMComputer - Issue with Get-CimInstance") }
						IF ($ErrorProcessInvokeCim) { Write-Warning -Message (Get-DefaultMessage -Message "$CIMComputer - Issue with Invoke-CimMethod") }
						Write-Warning -Message $Error[0].Exception.Message
					} #CATCH
					FINALLY
					{
						$CIMSplatting.Clear()
						$CIMInvokeSplatting.Clear()
					}
				}
			} #FOREACH ($Cim in $CimSessions)
		} #IF ($PSBoundParameters['CimSession'])
		ELSE
		{
			FOREACH ($Computer in $ComputerName)
			{
				$Computer = $Computer.ToUpper()
				
				IF ($PSCmdlet.ShouldProcess($Computer, "Disable Remote Desktop via Win32_TerminalServiceSetting"))
				{
					
					TRY
					{
						Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Test-Connection")
						IF (Test-Connection -Computer $Computer -count 1 -quiet)
						{
							$Splatting = @{
								Class = "Win32_TerminalServiceSetting"
								NameSpace = "root\cimv2\terminalservices"
								ComputerName = $Computer
								Authentication = 'PacketPrivacy'
								ErrorAction = 'Stop'
								ErrorVariable = 'ErrorProcessGetWmi'
							}
							
							IF ($PSBoundParameters['Credential'])
							{
								$Splatting.credential = $Credential
							}
							
							# disable Remote Desktop
							Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Get-WmiObject - disable Remote Desktop")
							(Get-WmiObject @Splatting).SetAllowTsConnections(0, 0) | Out-Null
							
							# Disable requirement that user must be authenticated
							#(Get-WmiObject -Class Win32_TSGeneralSetting @Splatting -Filter TerminalName='RDP-tcp').SetUserAuthenticationRequired(0)  Out-Null
						}
					}
					CATCH
					{
						Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Something wrong happened")
						IF ($ErrorProcessGetWmi) { Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Issue with Get-WmiObject") }
						Write-Warning -MEssage $Error[0].Exception.Message
					}
					FINALLY
					{
						$Splatting.Clear()
					}
				}
			} #FOREACH
		} #ELSE (Not CIM)
	} #PROCESS
} #Function
## End Disable-RemoteDesktop
## Begin Enable-RemoteDesktop
Function Enable-RemoteDesktop{
<#
	.SYNOPSIS
		The Function Enable-RemoteDesktop will enable RemoteDesktop on a local or remote machine.
	
	.DESCRIPTION
		The Function Enable-RemoteDesktop will enable RemoteDesktop on a local or remote machine.
	
	.PARAMETER ComputerName
		Specifies the computername
	
	.PARAMETER Credential
		Specifies the credential to use
	
	.PARAMETER CimSession
		Specifies one or more existing CIM Session(s) to use
	
	.EXAMPLE
		PS C:\> Enable-RemoteDesktop -ComputerName DC01
	
	.EXAMPLE
		PS C:\> Enable-RemoteDesktop -ComputerName DC01 -Credential (Get-Credential -cred "FX\SuperAdmin")
	
	.EXAMPLE
		PS C:\> Enable-RemoteDesktop -CimSession $Session
	
	.EXAMPLE
		PS C:\> Enable-RemoteDesktop -CimSession $Session1,$session2,$session3
	
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		www.lazywinadmin.com
		github.com/lazywinadmin
#>
	#Requires -RunAsAdministrator
	[CmdletBinding(DefaultParameterSetName = 'CimSession',
				   SupportsShouldProcess = $true)]
	param
	(
		[Parameter(ParameterSetName = 'Main',
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[Alias('CN', '__SERVER', 'PSComputerName')]
		[String[]]$ComputerName,
		
		[Parameter(ParameterSetName = 'Main')]
		[System.Management.Automation.Credential()]
		[Alias('RunAs')]
		$Credential = [System.Management.Automation.PSCredential]::Empty,
		
		[Parameter(ParameterSetName = 'CimSession')]
		[Microsoft.Management.Infrastructure.CimSession[]]$CimSession
	)
	
	BEGIN
	{
		# Helper Function
		Function Get-DefaultMessage
		{
<#
.SYNOPSIS
	Helper Function to show default message used in VERBOSE/DEBUG/WARNING
.DESCRIPTION
	Helper Function to show default message used in VERBOSE/DEBUG/WARNING
	and... HOST in some case.
	This is helpful to standardize the output messages
	
.PARAMETER Message
	Specifies the message to show
.NOTES
	Francois-Xavier Cat
	www.lazywinadmin.com
	@lazywinadm
#>
			PARAM ($Message)
			$DateFormat = Get-Date -Format 'yyyy/MM/dd-HH:mm:ss:ff'
			$FunctionName = (Get-Variable -Scope 1 -Name MyInvocation -ValueOnly).MyCommand.Name
			Write-Output "[$DateFormat][$FunctionName] $Message"
		} #Get-DefaultMessage
	}
	PROCESS
	{
		IF ($PSBoundParameters['CimSession'])
		{
			FOREACH ($Cim in $CimSession)
			{
				$CIMComputer = $($Cim.ComputerName).ToUpper()
				
				IF ($PSCmdlet.ShouldProcess($CIMComputer, "Enable Remote Desktop via Win32_TerminalServiceSetting"))
				{
					
					TRY
					{
						# Parameters for Get-CimInstance
						$CIMSplatting = @{
							Class = "Win32_TerminalServiceSetting"
							NameSpace = "root\cimv2\terminalservices"
							CimSession = $Cim
							Authentication = 'PacketPrivacy'
							ErrorAction = 'Stop'
							ErrorVariable = "ErrorProcessGetCimInstance"
						}
						
						# Parameters for Invoke-CimMethod
						$CIMInvokeSplatting = @{
							MethodName = "SetAllowTSConnections"
							Arguments = @{
								AllowTSConnections = 1
								ModifyFirewallException = 1
							}
							ErrorAction = 'Stop'
							ErrorVariable = "ErrorProcessInvokeCim"
						}
						
						Write-Verbose -Message (Get-DefaultMessage -Message "$CIMComputer - CIMSession - Enable Remote Desktop (and Modify Firewall Exception")
						Get-CimInstance @CIMSplatting | Invoke-CimMethod @CIMInvokeSplatting
					}
					CATCH
					{
						Write-Warning -Message (Get-DefaultMessage -Message "$CIMComputer - CIMSession - Something wrong happened")
						IF ($ErrorProcessGetCimInstance) { Write-Warning -Message (Get-DefaultMessage -Message "$CIMComputer - Issue with Get-CimInstance") }
						IF ($ErrorProcessInvokeCim) { Write-Warning -Message (Get-DefaultMessage -Message "$CIMComputer - Issue with Invoke-CimMethod") }
						Write-Warning -Message $Error[0].Exception.Message
					} #CATCH
					FINALLY
					{
						$CIMSplatting.Clear()
						$CIMInvokeSplatting.Clear()
					} #FINALLY
				} #$PSCmdlet.ShouldProcess
			} #FOREACH ($Cim in $CimSessions)
		} #IF ($PSBoundParameters['CimSession'])
		ELSE
		{
			FOREACH ($Computer in $ComputerName)
			{
				$Computer = $Computer.ToUpper()
				
				IF ($PSCmdlet.ShouldProcess($Computer, "Enable Remote Desktop via Win32_TerminalServiceSetting"))
				{
					TRY
					{
						Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Test-Connection")
						IF (Test-Connection -Computer $Computer -count 1 -quiet)
						{
							$Splatting = @{
								Class = "Win32_TerminalServiceSetting"
								NameSpace = "root\cimv2\terminalservices"
								ComputerName = $Computer
								Authentication = 'PacketPrivacy'
								ErrorAction = 'Stop'
								ErrorVariable = 'ErrorProcessGetWmi'
							}
							
							IF ($PSBoundParameters['Credential'])
							{
								$Splatting.credential = $Credential
							}
							
							# Enable Remote Desktop
							Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Get-WmiObject - Enable Remote Desktop")
							(Get-WmiObject @Splatting).SetAllowTsConnections(1, 1) | Out-Null
							
							# Disable requirement that user must be authenticated
							#(Get-WmiObject -Class Win32_TSGeneralSetting @Splatting -Filter TerminalName='RDP-tcp').SetUserAuthenticationRequired(0)  Out-Null
						}
					} #TRY
					CATCH
					{
						Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Something wrong happened")
						IF ($ErrorProcessGetWmi) { Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Issue with Get-WmiObject") }
						Write-Warning -MEssage $Error[0].Exception.Message
					} #CATCH
					FINALLY
					{
						$Splatting.Clear()
					} #FINALLY
				} #$PSCmdlet.ShouldProcess
			} #FOREACH
		} #ELSE (Not CIM)
	} #PROCESS
}
## End Enable-RemoteDesktop
## Begin RDP
Function RDP {
  <# 
  .SYNOPSIS 
  Remote Desktop Protocol to specified workstation(s) 

  .EXAMPLE 
  RDP Computer123456 

  .EXAMPLE 
  RDP 123456 
  #> 
	param(
	[Parameter(Mandatory=$true)]
	[string]$computername)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

	#Start Remote Desktop Protocol on specifed workstation
	& "C:\windows\system32\mstsc.exe" /v:$computername /fullscreen
}
## End RDP
## Begin Test-RemoteDesktopIsEnabled
Function Test-RemoteDesktopIsEnabled{
<#
.SYNOPSIS
  Function to check if RDP is enabled

.DESCRIPTION
  Function to check if RDP is enabled

.EXAMPLE
  Test-RemoteDesktopIsEnabled

  Test if Remote Desktop is enabled on the current machine

.EXAMPLE
  Test-RemoteDesktopIsEnabled -ComputerName SERVER01,SERVER02

  Test if Remote Desktop is enabled on the remote machine SERVER01 and SERVER02

.NOTES
	Francois-Xavier Cat
	@lazywinadm
	www.lazywinadmin.com
	github.com/lazywinadmin
#>


PARAM(
  [String[]]$ComputerName = $env:COMPUTERNAME
  )
  FOREACH ($Computer in $ComputerName)
  {
    TRY{
      IF (Test-Connection -Computer $Computer -count 1 -quiet)
      {
        $Splatting = @{
          ComputerName = $Computer
          NameSpace = "root\cimv2\TerminalServices"
        }
        # Enable Remote Desktop
        [boolean](Get-WmiObject -Class Win32_TerminalServiceSetting @Splatting).AllowTsConnections
        
        # Disable requirement that user must be authenticated
        #(Get-WmiObject -Class Win32_TSGeneralSetting @Splatting -Filter "TerminalName='RDP-tcp'").SetUserAuthenticationRequired(0) | Out-Null
      }
    }
    CATCH{
      Write-Warning -Message "Something wrong happened"
      Write-Warning -MEssage $Error[0].Exception.Message
    }
  }#FOREACH
  
}
## End Test-RemoteDesktopIsEnabled