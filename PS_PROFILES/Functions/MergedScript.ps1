Function ConnectTo-Database{
    # Begin Connection Config
    $SQLServer = "LASPSHOST.Contoso.CORP"
    $SQLDB = "SITEST"
    # Credentials for sa user
    $SQLCred = Import-Clixml -Path .\Creds\SQLCredsSA.xml
    # End Connection Config

    # Begin Region Connection
    Import-Module dbatools
    Write-Host "Building connection string" -ForegroundColor Black -BackgroundColor White
    $Connection = Connect-DbaInstance -SqlInstance $SqlServer -Credential $SQLCred
    Write-Host "Opening connection to $($SqlServer)"
    # End Region Connection

}

Function Convert-ChassisType {
    Param ([int[]]$ChassisType)
    $List = New-Object System.Collections.ArrayList
    Switch ($ChassisType) {
        0x0001  {[void]$List.Add('Other')}
        0x0002  {[void]$List.Add('Unknown')}
        0x0003  {[void]$List.Add('Desktop')}
        0x0004  {[void]$List.Add('Low Profile Desktop')}
        0x0005  {[void]$List.Add('Pizza Box')}
        0x0006  {[void]$List.Add('Mini Tower')}
        0x0007  {[void]$List.Add('Tower')}
        0x0008  {[void]$List.Add('Portable')}
        0x0009  {[void]$List.Add('Laptop')}
        0x000A  {[void]$List.Add('Notebook')}
        0x000B  {[void]$List.Add('Hand Held')}
        0x000C  {[void]$List.Add('Docking Station')}
        0x000D  {[void]$List.Add('All in One')}
        0x000E  {[void]$List.Add('Sub Notebook')}
        0x000F  {[void]$List.Add('Space-Saving')}
        0x0010  {[void]$List.Add('Lunch Box')}
        0x0011  {[void]$List.Add('Main System Chassis')}
        0x0012  {[void]$List.Add('Expansion Chassis')}
        0x0013  {[void]$List.Add('Subchassis')}
        0x0014  {[void]$List.Add('Bus Expansion Chassis')}
        0x0015  {[void]$List.Add('Peripheral Chassis')}
        0x0016  {[void]$List.Add('Storage Chassis')}
        0x0017  {[void]$List.Add('Rack Mount Chassis')}
        0x0018  {[void]$List.Add('Sealed-Case PC')}
    }
    $List -join ', '
}

Function ConvertFrom-Base64
{
	<#
	.SYNOPSIS
		Converts the specified string, which encodes binary data as base-64 digits, to an equivalent 8-bit unsigned integer array.
	
	.DESCRIPTION
		Converts the specified string, which encodes binary data as base-64 digits, to an equivalent 8-bit unsigned integer array.
	
	.PARAMETER String
		Specifies the String to Convert
		
	.EXAMPLE
		ConvertFrom-Base64 -String $ImageBase64 |Out-File ImageTest.png
	
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		www.lazywinadmin.com
		github.com/lazywinadmin
#>
	[CmdletBinding()]
	PARAM (
		[parameter(Mandatory = $true, ValueFromPipeline)]
		[String]$String
	)
	TRY
	{
		Write-Verbose -Message "[ConvertFrom-Base64] Converting String"
		[System.Text.Encoding]::Default.GetString(
		[System.Convert]::FromBase64String($String)
		)
	}
	CATCH
	{
		Write-Error -Message "[ConvertFrom-Base64] Something wrong happened"
		$Error[0].Exception.Message
	}
}

Function ConvertTo-Base64
{
<#
	.SYNOPSIS
		Function to convert an image to Base64
	
	.DESCRIPTION
		Function to convert an image to Base64
	
	.PARAMETER Path
		Specifies the path of the file
	
	.EXAMPLE
		ConvertTo-Base64 -Path "C:\images\PowerShellLogo.png"
	
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		www.lazywinadmin.com
		github.com/lazywinadmin
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[ValidateScript({ Test-Path -Path $_ })]
		[String]$Path
	)
	Write-Verbose -Message "[ConvertTo-Base64] Converting image to Base64 $Path"
	[System.convert]::ToBase64String((Get-Content -Path $path -Encoding Byte))
}

Function ConvertTo-StringList
{
<#
	.SYNOPSIS
		Function to convert an array into a string list with a delimiter.
	
	.DESCRIPTION
		Function to convert an array into a string list with a delimiter.
	
	.PARAMETER Array
		Specifies the array to process.
	
	.PARAMETER Delimiter
		Separator between value, default is ","
	
	.EXAMPLE
		$Computers = "Computer1","Computer2"
		ConvertTo-StringList -Array $Computers
	
		Output: 
		Computer1,Computer2
	
	.EXAMPLE
		$Computers = "Computer1","Computer2"
		ConvertTo-StringList -Array $Computers -Delimiter "__"
	
		Output: 
		Computer1__Computer2
	
	.EXAMPLE
		$Computers = "Computer1"
		ConvertTo-StringList -Array $Computers -Delimiter "__"
	
		Output: 
		Computer1
		
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
	
		I used this Function in System Center Orchestrator (SCORCH).
		This is sometime easier to pass data between activities
#>
	
	[CmdletBinding()]
	[OutputType([string])]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[System.Array]$Array,
		
		[system.string]$Delimiter = ","
	)
	
	BEGIN { $StringList = "" }
	PROCESS
	{
		Write-Verbose -Message "Array: $Array"
		foreach ($item in $Array)
		{
			# Adding the current object to the list
			$StringList += "$item$Delimiter"
		}
		Write-Verbose "StringList: $StringList"
	}
	END
	{
		TRY
		{
			IF ($StringList)
			{
				$lenght = $StringList.Length
				Write-Verbose -Message "StringList Lenght: $lenght"
				
				# Output Info without the last delimiter
				$StringList.Substring(0, ($lenght - $($Delimiter.length)))
			}
		}# TRY
		CATCH
		{
			Write-Warning -Message "[END] Something wrong happening when output the result"
			$Error[0].Exception.Message
		}
		FINALLY
		{
			# Reset Variable
			$StringList = ""
		}
	}
}

Function Copy-Folder([string]$source, [string]$destination, [bool]$recursive) {
    if (!$(Test-Path($destination))) {
        New-Item $destination -type directory -Force
    }
####################################################################################################
# This Function copies a folder (and optionally, its subfolders)
#
# When copying subfolders it calls itself recursively
#
# Requires WebClient object $webClient defined, e.g. $webClient = New-Object System.Net.WebClient
#
# Parameters:
#   $source      - The url of folder to copy, with trailing /, e.g. http://website/folder/structure/
#   $destination - The folder to copy $source to, with trailing \ e.g. D:\CopyOfStructure\
#   $recursive   - True if subfolders of $source are also to be copied or False to ignore subfolders
#   Return       - None
####################################################################################################
    # Get the file list from the web page
    $webString = $webClient.DownloadString($source)
    $lines = [Regex]::Split($webString, "<br>")
    # Parse each line, looking for files and folders
    foreach ($line in $lines) {
        if ($line.ToUpper().Contains("HREF")) {
            # File or Folder
            if (!$line.ToUpper().Contains("[TO PARENT DIRECTORY]")) {
                # Not Parent Folder entry
                $items =[Regex]::Split($line, """")
                $items = [Regex]::Split($items[2], "(>|<)")
                $item = $items[2]
                if ($line.ToLower().Contains("&lt;dir&gt")) {
                    # Folder
                    if ($recursive) {
                        # Subfolder copy required
                        Copy-Folder "$source$item/" "$destination$item/" $recursive
                    } else {
                        # Subfolder copy not required
                    }
                } else {
                    # File
                    $webClient.DownloadFile("$source$item", "$destination$item")
                }
            }
        }
    }
}

Function Create-HTMLTable
{
param([array]$Array)
$arrHTML = $Array | ConvertTo-Html
$arrHTML[-1] = $arrHTML[-1].ToString().Replace(‘</body></html>’,"")
Return $arrHTML[5..2000]
}


Function Create-WMIFilters 
Import-Module ActiveDirectory     # Get-Help *AD* 
{ 
    # Importing or adding a WMI Filter object into AD is a system only operation.  
    # You need to enable system only changes on a domain controller for a successful import.  
    # To do this, on the domain controller you are using for importing, open the registry editor and create the following registry value. 
    # 
    # Key: HKLM\System\CurrentControlSet\Services\NTDS\Parameters  
    # Value Name: Allow System Only Change  
    # Value Type: REG_DWORD  
    # Value Data: 1 (Binary) 
    # 
    # Put this somewhere in your master code: new-itemproperty "HKLM:\System\CurrentControlSet\Services\NTDS\Parameters" -name "Allow System Only Change" -value 1 -propertyType dword 
 
 
    # Name,Query,Description 
    $WMIFilters = @(   ('Virtual Machines', 'SELECT * FROM Win32_ComputerSystem WHERE Model = "Virtual Machine"', 'Hyper-V'), 
                    ('Workstation 32-bit', 'Select * from WIN32_OperatingSystem where ProductType=1 Select * from Win32_Processor where AddressWidth = "32"', ''), 
                    ('Workstation 64-bit', 'Select * from WIN32_OperatingSystem where ProductType=1 Select * from Win32_Processor where AddressWidth = "64"', ''), 
                    ('Workstations', 'SELECT * FROM Win32_OperatingSystem WHERE ProductType = "1"', ''), 
                    ('Domain Controllers', 'SELECT * FROM Win32_OperatingSystem WHERE ProductType = "2"', ''), 
                    ('Servers', 'SELECT * FROM Win32_OperatingSystem WHERE ProductType = "3"', ''), 
                    ('Resources', 'Select * from Win32_LogicalDisk where FreeSpace > 629145600 AND Description <> "Network Connection"', 'Target only machines that have at least 600 megabytes (MB) available.'), 
                    ('Hotfix', 'Select * from Win32_QuickFixEngineering where HotFixID = "q147222"', 'Apply a policy on computers that have a specific hotfix.'), 
                    ('Time zone', 'Select * from win32_timezone where bias =120', 'Apply policy on all servers located in the South of Africa.'), 
                    ('Configuration', 'Select * from Win32_NetworkProtocol where SupportsMulticasting = true', 'Avoid turning on netmon on computers that can have multicasting turned on.'), 
                    ('Windows 2000', 'select * from Win32_OperatingSystem where Version like "5.0%"', 'This is used to filter out GPOs that are only meant for Windows 2000 systems and should not apply to newer OSes eventhough Windows 2000 does not support WMI filtering'), 
                    ('Windows XP', 'select * from Win32_OperatingSystem where (Version like "5.1%" or Version like "5.2%") and ProductType = "1"', ''), 
                    ('Windows Vista', 'select * from Win32_OperatingSystem where Version like "6.0%" and ProductType = "1"', ''), 
                    ('Windows 7', 'select * from Win32_OperatingSystem where Version like "6.1%" and ProductType = "1"', ''), 
                    ('Windows Server 2003', 'select * from Win32_OperatingSystem where Version like "5.2%" and ProductType = "3"', ''), 
                    ('Windows Server 2008', 'select * from Win32_OperatingSystem where Version like "6.0%" and ProductType = "3"', ''), 
                    ('Windows Server 2008 R2', 'select * from Win32_OperatingSystem where Version like "6.1%" and ProductType = "3"', ''), 
                    ('Windows Vista and Windows Server 2008', 'select * from Win32_OperatingSystem where Version like "6.0%" and ProductType<>"2"', ''), 
                    ('Windows Server 2003 and Windows Server 2008', 'select * from Win32_OperatingSystem where (Version like "5.2%" or Version like "6.0%") and ProductType="3"', ''), 
                    ('Windows 2000, XP and 2003', 'select * from Win32_OperatingSystem where Version like "5.%" and ProductType<>"2"', ''), 
                    ('Manufacturer Dell', 'Select * from WIN32_ComputerSystem where Manufacturer = "DELL"', ''), 
                    ('Installed Memory > 1Gb', 'Select * from WIN32_ComputerSystem where TotalPhysicalMemory >= 1073741824', '') 
                ) 
 
    $defaultNamingContext = (get-adrootdse).defaultnamingcontext  
    $configurationNamingContext = (get-adrootdse).configurationNamingContext  
    $msWMIAuthor = "Administrator@" + [System.DirectoryServices.ActiveDirectory.Domain]::getcurrentdomain().name 
     
    for ($i = 0; $i -lt $WMIFilters.Count; $i++)  
    { 
        $WMIGUID = [string]"{"+([System.Guid]::NewGuid())+"}"    
        $WMIDN = "CN="+$WMIGUID+",CN=SOM,CN=WMIPolicy,CN=System,"+$defaultNamingContext 
        $WMICN = $WMIGUID 
        $WMIdistinguishedname = $WMIDN 
        $WMIID = $WMIGUID 
 
        $now = (Get-Date).ToUniversalTime() 
        $msWMICreationDate = ($now.Year).ToString("0000") + ($now.Month).ToString("00") + ($now.Day).ToString("00") + ($now.Hour).ToString("00") + ($now.Minute).ToString("00") + ($now.Second).ToString("00") + "." + ($now.Millisecond * 1000).ToString("000000") + "-000" 
 
        $msWMIName = $WMIFilters[$i][0] 
        $msWMIParm1 = $WMIFilters[$i][2] + " " 
        $msWMIParm2 = "1;3;10;" + $WMIFilters[$i][1].Length.ToString() + ";WQL;root\CIMv2;" + $WMIFilters[$i][1] + ";" 
 
        $Attr = @{"msWMI-Name" = $msWMIName;"msWMI-Parm1" = $msWMIParm1;"msWMI-Parm2" = $msWMIParm2;"msWMI-Author" = $msWMIAuthor;"msWMI-ID"=$WMIID;"instanceType" = 4;"showInAdvancedViewOnly" = "TRUE";"distinguishedname" = $WMIdistinguishedname;"msWMI-ChangeDate" = $msWMICreationDate; "msWMI-CreationDate" = $msWMICreationDate} 
        $WMIPath = ("CN=SOM,CN=WMIPolicy,CN=System,"+$defaultNamingContext) 
     
        New-ADObject -name $WMICN -type "msWMI-Som" -Path $WMIPath -OtherAttributes $Attr 
    } 
 
} 
## Begin Disable-RemoteDesktop
Function Disable-RemoteDesktop
{
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
## Begin Disconnect-ViSession
Function Disconnect-ViSession {
<#
.SYNOPSIS
Disconnects a connected vCenter Session.

.DESCRIPTION
Disconnects a open connected vCenter Session.

.PARAMETER  SessionList
A session or a list of sessions to disconnect.

.EXAMPLE
PS C:\> Get-VISession | Where { $_.IdleMinutes -gt 5 } | Disconnect-ViSession

.EXAMPLE
PS C:\> Get-VISession | Where { $_.Username -eq “User19” } | Disconnect-ViSession
#>
[CmdletBinding()]
Param (
[Parameter(ValueFromPipeline=$true)]
$SessionList
)
Process {
$SessionMgr = Get-View $DefaultViserver.ExtensionData.Client.ServiceContent.SessionManager
$SessionList | Foreach {
Write “Disconnecting Session for $($_.Username) which has been active since $($_.LoginTime)”
$SessionMgr.TerminateSession($_.Key)
}
}
}
## End Disconnect-ViSession
## Begin Enable-RemoteDesktop
Function Enable-RemoteDesktop
{
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
## Begin Expand-ScriptAlias
Function Expand-ScriptAlias
{
<#
	.SYNOPSIS
		Function to replace Aliases used in a script by their fullname
	
	.DESCRIPTION
		Function to replace Aliases used in a script by their fullname.
		Using PowerShell AST we are able to retrieve the Functions and cmdlets used in a script.
	
	.PARAMETER Path
		Specifies the Path to the file.
		Alias: FullName
	
	.EXAMPLE
		"C:\LazyWinAdmin\testscript.ps1", "C:\LazyWinAdmin\testscript2.ps1" | Expand-ScriptAlias
	
	.EXAMPLE
		gci C:\LazyWinAdmin -File | Expand-ScriptAlias
	
	.EXAMPLE
		Expand-ScriptAlias -Path "C:\LazyWinAdmin\testscript.ps1"

    .EXAMPLE
        "C:\LazyWinAdmin\testscript.ps1", "C:\LazyWinAdmin\testscript2.ps1" | Expand-ScriptAlias -Confirm

    .EXAMPLE
        "C:\LazyWinAdmin\testscript.ps1", "C:\LazyWinAdmin\testscript2.ps1" | Expand-ScriptAlias -WhatIf

        What if: Performing the operation "Expand Alias: select to Select-Object (startoffset: 15)" on target "C:\LazyWinAdmin\testscript2.ps1".
        What if: Performing the operation "Expand Alias: sort to Sort-Object (startoffset: 10)" on target "C:\LazyWinAdmin\testscript2.ps1".
        What if: Performing the operation "Expand Alias: group to Group-Object (startoffset: 4)" on target "C:\LazyWinAdmin\testscript2.ps1".
        What if: Performing the operation "Expand Alias: gci to Get-ChildItem (startoffset: 0)" on target "C:\LazyWinAdmin\testscript2.ps1".
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
	PARAM (
		[Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[ValidateScript({ Test-Path -Path $_ })]
		[Alias('FullName')]
		[System.String]$Path
	)
	PROCESS
	{
		FOREACH ($File in $Path)
		{
			Write-Verbose -Message '[PROCESS] $File'
			
			TRY
			{
				# Retrieve file content
				$ScriptContent = (Get-Content $File -Delimiter $([char]0))
				
				# AST Parsing
				$AbstractSyntaxTree = [System.Management.Automation.Language.Parser]::
				ParseInput($ScriptContent, [ref]$null, [ref]$null)
				
				# Find Aliases
				$Aliases = $AbstractSyntaxTree.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true) |
				ForEach-Object -Process {
					$Command = $_.CommandElements[0]
					if ($Alias = Get-Alias | Where-Object { $_.Name -eq $Command })
					{
						
						# Output information
						[PSCustomObject]@{
							File = $File
							Alias = $Alias.Name
							Definition = $Alias.Definition
							StartLineNumber = $Command.Extent.StartLineNumber
							EndLineNumber = $Command.Extent.EndLineNumber
							StartColumnNumber = $Command.Extent.StartColumnNumber
							EndColumnNumber = $Command.Extent.EndColumnNumber
							StartOffset = $Command.Extent.StartOffset
							EndOffset = $Command.Extent.EndOffset
							
						}#[PSCustomObject]
					}#if ($Alias)
				} | Sort-Object -Property EndOffset -Descending
				
				# The sort-object is important, we change the values from the end first to not lose the positions of every aliases.
				Foreach ($Alias in $Aliases)
				{
					# whatif and confirm support
					if ($psCmdlet.ShouldProcess($file, "Expand Alias: $($Alias.alias) to $($Alias.definition) (startoffset: $($alias.StartOffset))"))
					{
						# Remove alias and insert full cmldet name
						$ScriptContent = $ScriptContent.Remove($Alias.StartOffset, ($Alias.EndOffset - $Alias.StartOffset)).Insert($Alias.StartOffset, $Alias.Definition)
						# Apply to the file
						Set-Content -Path $File -Value $ScriptContent -Confirm:$false
					}
				}#ForEach Alias in Aliases
				
			}#TRY
			CATCH
			{
				Write-Error -Message $($Error[0].Exception.Message)
			}
		}#FOREACH File in Path
	}#PROCESS
}#Expand-ScriptAlias
## End Expand-ScriptAlias
## Begin Export-Xls 
Function Export-Xls 
{ 
 
<# 
.SYNOPSIS 
Export to Excel file. 
 
.DESCRIPTION 
Export to Excel file. Since Excel files can have multiple worksheets, you can specify the name of the Excel file and worksheet. Exports to a worksheet named "Sheet" by default. 
 
.PARAMETER Path 
Specifies the path to the Excel file to export. 
Note: The path must contain an extension for spreadsheets, such as .xls, .xlsx, .xlsm, .xml, and .ods 
 
.PARAMETER Worksheet 
Specifies the name of the worksheet where the data is exported. The default is "Sheet". 
Note: If a worksheet already exists with the given name, no error occurs. The name will be appended with (2), or (3), or (4), etc. 
 
.PARAMETER InputObject 
Specifies the objects to export. You can also pipe objects to Export-Xls. 
 
.PARAMETER Append 
Append the exported data to a new worksheet in the excel file. 
If you Append to a spreadsheet that does not allow more than one worksheet, the new data will not be saved. 
Note: For this Function, -Append is not considered clobbering the file, but modifying the file, so -Append and -NoClobber do not conflict with each other. 
 
.PARAMETER NoClobber 
Do not overwrite the file. 
Use -Append if you want to add a worksheet to the excel file, but leave the others intact. 
Note: For this Function, -Append is not considered clobbering the file, but modifying the file, so -Append and -NoClobber do not conflict with each other. 
 
.PARAMETER NoTypeInformation 
Omits the type information. 
 
.INPUTS 
System.Management.Automation.PSObject 
 
.OUTPUTS 
System.String 
This is a CSV list, which is then exported to a csv file, which is then converted to an Excel file. 
 
.EXAMPLE 
Get-Process | Export-Xls ".\export.xlsx" -Worksheet "Sheet1" 
Export the output of Get-Process to Worksheet "Sheet1" of export.xlsx 
Note: export.xlsx is overwritten. 
 
.EXAMPLE 
Get-Process | Export-Xls ".\export.xlsx" -Worksheet "Sheet2" -NoTypeInformation 
Export the output of Get-Process to Worksheet "Sheet2" of export.xlsx with no type information 
Note: export.xlsx is overwritten. 
 
.EXAMPLE 
Get-Process | Export-Xls ".\export.xlsx" -Worksheet "Sheet3" -Append 
Export output of Get-Process to Worksheet "Sheet3" and Append it to export.xlsx 
Note: export.xlsx is modified. 
 
.EXAMPLE 
Get-Process | Export-Xls ".\export.xlsx" -Worksheet "Sheet4" -NoClobber 
Export output of Get-Process to Worksheet "Sheet4" and create export.xlsx if it doesn't exist. 
Note: export.xlsx is created. If export.xlsx already exist, the Function terminates with an error. 
 
.EXAMPLE 
(Get-Alias s*), (Get-Alias g*) | Export-Xls ".\export.xlsx" -Worksheet "Alias" 
Export Aliases that start with s and g to Worksheet "Alias" of export.xlsx 
Note: See next example for possible problems when doing something like this 
 
.EXAMPLE 
(Get-Alias), (Get-Process) | Export-Xls ".\export.xlsx" -Worksheet "Alias and Process" 
Export the result of Get-Command and Get-Process to Worksheet "Alias and Process" of export.xlsx 
Note: Since Get-Alias and Get-Process do not return objects with the same properties, not all information is recorded. 
 
.LINK 
Export-Xls 
http://gallery.technet.microsoft.com/scriptcenter/d41565f1-37ef-43cb-9462-a08cd5a610e2 
Import-Xls 
http://gallery.technet.microsoft.com/scriptcenter/17bcabe7-322a-43d3-9a27-f3f96618c74b 
Import-Csv 
Export-Csv 
 
.NOTES 
Author: Francis de la Cerna 
Created: 2011-03-27 
Modified: 2011-04-09 
#Requires –Version 2.0 
#> 
 
    [CmdletBinding(SupportsShouldProcess=$true)] 
 
    Param( 
        [parameter(mandatory=$true, position=1)] 
        $Path, 
     
        [parameter(mandatory=$false, position=2)] 
        $Worksheet = "Sheet", 
     
        [parameter( 
            mandatory=$true,  
            ValueFromPipeline=$true, 
            ValueFromPipelineByPropertyName=$true)] 
        [psobject[]] 
        $InputObject, 
     
        [parameter(mandatory=$false)] 
        [switch] 
        $Append, 
     
        [parameter(mandatory=$false)] 
        [switch] 
        $NoClobber, 
 
        [parameter(mandatory=$false)] 
        [switch] 
        $NoTypeInformation, 
 
        [parameter(mandatory=$false)] 
        [switch] 
        $Force 
    ) 
     
    Begin 
    { 
        # WhatIf, Confirm, Verbose 
        # Probably not the way to do it, but this Function runs all or nothing 
        # so, exit each block (Begin, Process, End) if shouldProcesss is false. 
        # Disabled confirmations on operations on temporary files, but enabled 
        # verbose messages. 
        #  
        $shouldProcess = $Force -or $psCmdlet.ShouldProcess($Path); 
         
        if (-not $shouldProcess) { return; } 
         
        Function GetTempFileName($extension) 
        { 
            $temp = [io.path]::GetTempFileName(); 
            $params = @{ 
                Path = $temp; 
                Destination = $temp + $extension; 
                Confirm = $false; 
                Verbose = $VerbosePreference; 
            } 
            Move-Item @params; 
            $temp += $extension; 
            return $temp; 
        } 
         
        # check extension of $Path to see what excel format to export to 
        # since an extension like .xls can have multiple formats, this 
        # will need to be changed 
        # 
        $xlFileFormats = @{ 
            # single worksheet formats 
            '.csv'  = 6;        # 6, 22, 23, 24 
            '.dbf'  = 11;       # 7, 8, 11 
            '.dif'  = 9;        #  
            '.prn'  = 36;       #  
            '.slk'  = 2;        # 2, 10 
            '.wk1'  = 31;       # 5, 30, 31 
            '.wk3'  = 32;       # 15, 32 
            '.wk4'  = 38;       #  
            '.wks'  = 4;        #  
            '.xlw'  = 35;       #  
             
            # multiple worksheet formats 
            '.xls'  = -4143;    # -4143, 1, 16, 18, 29, 33, 39, 43 
            '.xlsb' = 50;       # 
            '.xlsm' = 52;       # 
            '.xlsx' = 51;       # 
            '.xml'  = 46;       # 
            '.ods'  = 60;       # 
        } 
         
        $ext = [io.path]::GetExtension($Path).toLower(); 
        if ($xlFileFormats.Keys -notcontains $ext) { 
            $msg = "Error: $Path has unknown extension. Try "; 
            foreach ($extension in ($xlFileFormats.Keys | sort)) { 
                $msg += "$extension "; 
            } 
            Throw "$msg"; 
        } 
         
        # get full path 
        # 
        if (-not [io.path]::IsPathRooted($Path)) { 
            $fswd = $psCmdlet.CurrentProviderLocation("FileSystem"); 
            $Path = Join-Path -Path $fswd -ChildPath $Path; 
        } 
         
        $Path = [io.path]::GetFullPath($Path); 
 
        $obj = New-Object System.Collections.ArrayList; 
    } 
 
    Process 
    { 
        if (-not $shouldProcess) { return; } 
 
        $InputObject | ForEach-Object{ $obj.Add($_) | Out-Null; } 
    } 
 
    End 
    { 
        if (-not $shouldProcess) { return; } 
         
        $xl = New-Object -ComObject Excel.Application; 
        $xl.DisplayAlerts = $false; 
        $xl.Visible = $false; 
         
        # create temporary .csv file from all $InputObject 
        # 
        $csvTemp = GetTempFileName(".csv"); 
        $obj | Export-Csv -Path $csvTemp -Force -NoType:$NoTypeInformation -Confirm:$false; 
         
        # create a temporary excel file from the temporary .csv file 
        # 
        $xlsTemp = GetTempFileName($ext); 
        $wb = $xl.Workbooks.Add($csvTemp); 
        $ws = $wb.Worksheets.Item(1); 
        $ws.Name = $Worksheet; 
        $wb.SaveAs($xlsTemp, $xlFileFormats[$ext]); 
        $xlsTempSaved = $?; 
        $wb.Close(); 
        Remove-Variable -Name ('ws', 'wb') -Confirm:$false; 
         
        if ($xlsTempSaved) { 
            # decide how to export based on switches and $Path 
            # 
            $fileExist = Test-Path $Path; 
            $createFile = -not $fileExist; 
            $appendFile = $fileExist -and $Append; 
            $clobberFile = $fileExist -and (-not $appendFile) -and (-not $NoClobber); 
            $needNewFile = $fileExist -and (-not $appendFile) -and $NoClobber; 
         
            if ($appendFile) { 
                $wbDst = $xl.Workbooks.Open($Path); 
                $wbSrc = $xl.Workbooks.Open($xlsTemp); 
                $wsDst = $wbDst.Worksheets.Item($wbDst.Worksheets.Count); 
                $wsSrc = $wbSrc.Worksheets.Item(1); 
                $wsSrc.Name = $Worksheet; 
                $wsSrc.Copy($wsDst); 
                $wsDst.Move($wbDst.Worksheets.Item($wbDst.Worksheets.Count-1)); 
                $wbDst.Worksheets.Item(1).Select(); 
                $wbSrc.Close($false); 
                $wbDst.Close($true); 
                Remove-Variable -Name ('wsSrc', 'wbSrc') -Confirm:$false; 
                Remove-Variable -Name ('wsDst', 'wbDst') -Confirm:$false; 
            } elseif ($createFile -or $clobberFile) { 
                Copy-Item $xlsTemp -Destination $Path -Force -Confirm:$false; 
            } elseif ($needNewFile) { 
                Write-Error "The file '$Path' already exists." -Category ResourceExists; 
            } else { 
                Write-Error "Something was wrong with my logic."; 
            } 
        } 
         
        # clean up 
        # 
        $xl.Quit(); 
        Remove-Variable -name xl -Confirm:$false; 
        Remove-Item $xlsTemp -Confirm:$false -Verbose:$VerbosePreference; 
        Remove-Item $csvTemp -Confirm:$false -Verbose:$VerbosePreference; 
        [gc]::Collect(); 
    } 
} 
## End Export-Xls 
## Begin FindWordDocumentEnd 
Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}
## End FindWordDocumentEnd 
## Begin Force-WSUSCheckin
Function Force-WSUSCheckin($Computer)
{
   Invoke-Command -computername $Computer -scriptblock { Start-Service wuauserv -Verbose }
   # Have to use psexec with the -s parameter as otherwise we receive an "Access denied" message loading the comobject
   $Cmd = '$updateSession = new-object -com "Microsoft.Update.Session";$updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates'
   psexec -s \\$Computer powershell.exe -command $Cmd
   Write-host "Waiting 10 seconds for SyncUpdates webservice to complete to add to the wuauserv queue so that it can be reported on"
   Start-sleep -seconds 10
   Invoke-Command -computername $Computer -scriptblock
   {
      # Now that the system is told it CAN report in, run every permutation of commands to actually trigger the report in operation
      wuauclt /detectnow
      (New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow()
      wuauclt /reportnow
      c:\windows\system32\UsoClient.exe startscan
   }
}
## End Force-WSUSCheckin
## Begin Functions-PSStoredCredentials - Should be module!
<#
.SYNOPSIS
Functions-PSStoredCredentials - PowerShell Functions to manage stored credentials for re-use

.DESCRIPTION 
This script adds two Functions that can be used to manage stored credentials
on your admin workstation.

.EXAMPLE
. .\Functions-PSStoredCredentials.ps1

.LINK
https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp
#>


Function New-StoredCredential {

    <#
    .SYNOPSIS
    New-StoredCredential - Create a new stored credential

    .DESCRIPTION 
    This Function will save a new stored credential to a .cred file.

    .EXAMPLE
    New-StoredCredential

    .LINK
    https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
    .NOTES
    Written by: Paul Cunningham

    Find me on:

    * My Blog:	http://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham
    * LinkedIn:	http://au.linkedin.com/in/cunninghamp/
    * Github:	https://github.com/cunninghamp

    For more Office 365 tips, tricks and news
    check out Practical 365.

    * Website:	https://practical365.com
    * Twitter:	https://twitter.com/practical365
    #>

    if (!(Test-Path Variable:\KeyPath)) {
        Write-Warning "The `$KeyPath variable has not been set. Consider adding `$KeyPath to your PowerShell profile to avoid this prompt."
        $path = Read-Host -Prompt "Enter a path for stored credentials"
        Set-Variable -Name KeyPath -Scope Global -Value $path

        if (!(Test-Path $KeyPath)) {
        
            try {
                New-Item -ItemType Directory -Path $KeyPath -ErrorAction STOP | Out-Null
            }
            catch {
                throw $_.Exception.Message
            }           
        }
    }

    $Credential = Get-Credential -Message "Enter a user name and password"

    $Credential.Password | ConvertFrom-SecureString | Out-File "$($KeyPath)\$($Credential.Username).cred" -Force

    # Return a PSCredential object (with no password) so the caller knows what credential username was entered for future recalls
    New-Object -TypeName System.Management.Automation.PSCredential($Credential.Username,(new-object System.Security.SecureString))

}



Function Get-StoredCredential {

    <#
    .SYNOPSIS
    Get-StoredCredential - Retrieve or list stored credentials

    .DESCRIPTION 
    This Function can be used to list available credentials on
    the computer, or to retrieve a credential for use in a script
    or command.

    .PARAMETER UserName
    Get the stored credential for the username

    .PARAMETER List
    List the stored credentials on the computer

    .EXAMPLE
    Get-StoredCredential -List

    .EXAMPLE
    $credential = Get-StoredCredential -UserName admin@tenant.onmicrosoft.com

    .EXAMPLE
    Get-StoredCredential -List

    .LINK
    https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
    .NOTES
    Written by: Paul Cunningham

    Find me on:

    * My Blog:	http://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham
    * LinkedIn:	http://au.linkedin.com/in/cunninghamp/
    * Github:	https://github.com/cunninghamp

    For more Office 365 tips, tricks and news
    check out Practical 365.

    * Website:	https://practical365.com
    * Twitter:	https://twitter.com/practical365
    #>

    param(
        [Parameter(Mandatory=$false, ParameterSetName="Get")]
        [string]$UserName,
        [Parameter(Mandatory=$false, ParameterSetName="List")]
        [switch]$List
        )

    if (!(Test-Path Variable:\KeyPath)) {
        Write-Warning "The `$KeyPath variable has not been set. Consider adding `$KeyPath to your PowerShell profile to avoid this prompt."
        $path = Read-Host -Prompt "Enter a path for stored credentials"
        Set-Variable -Name KeyPath -Scope Global -Value $path
    }


    if ($List) {

        try {
        $CredentialList = @(Get-ChildItem -Path $keypath -Filter *.cred -ErrorAction STOP)

        foreach ($Cred in $CredentialList) {
            Write-Host "Username: $($Cred.BaseName)"
            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }

    }

    if ($UserName) {
        if (Test-Path "$($KeyPath)\$($Username).cred") {
        
            $PwdSecureString = Get-Content "$($KeyPath)\$($Username).cred" | ConvertTo-SecureString
            
            $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $PwdSecureString
        }
        else {
            throw "Unable to locate a credential for $($Username)"
        }

        return $Credential
    }
}
## End Functions-PSStoredCredentials - Should be module!
## Begin Get-ActivationStatus - Needs to be fixed!
Function Get-ActivationStatus {

    $Servers = GC C:\LazyWinAdmin\Servers\Servers-All-Alive2.txt

    ForEach ($server in $Servers){
            $wpa = Get-CimInstance SoftwareLicensingProduct -ComputerName $server -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" -Property LicenseStatus -ErrorAction SilentlyContinue
        } catch {
            $status = New-Object ComponentModel.Win32Exception ($_.Exception.ErrorCode)
            $wpa = $null    
        }
        $out = New-Object psobject -Property @{
            ComputerName = $server;
            Status = [string]::Empty;
        }
        if ($wpa) {
            :outer foreach($item in $wpa) {
                switch ($item.LicenseStatus) {
                    0 {$out.Status = "Unlicensed"}
                    1 {$out.Status = "Licensed"; break outer}
                    2 {$out.Status = "Out-Of-Box Grace Period"; break outer}
                    3 {$out.Status = "Out-Of-Tolerance Grace Period"; break outer}
                    4 {$out.Status = "Non-Genuine Grace Period"; break outer}
                    5 {$out.Status = "Notification"; break outer}
                    6 {$out.Status = "Extended Grace"; break outer}
                    default {$out.Status = "Unknown value"}
                }
            }
        } else {$out.Status = $status.Message}
        $out
}
## End Get-ActivationStatus - Needs to be fixed!
## Begin Get-ADDirectReports
Function Get-ADDirectReports
{
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
Function Get-ADDomains
{
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
Function Get-ADSITokenGroup
{
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
## Begin Get-ADUserLastLogon
Function Get-ADUserLastLogon([string]$userName)
{
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
#Requires -Version 3.0

## End Get-ADUserLastLogon
## Begin Get-AntiSpyware
Function Get-AntiSpyware
{

<#
.Synopsis
   Get AntiSpyware information

.DESCRIPTION
   Obtain from cim (wmi) as SecutiryCenter shows.
   make sure your OS is Workstation, not as Server. (Because server does not have secutiry Center.)

.EXAMPLE
    # this will obtain from localhost
    Get-AntiSpyware

.EXAMPLE
    # this will obtain from 192.168.100.1 with credential you enter.
    $cred = Get-Credential
    Get-AntiSpyware -computerName 192.168.100.1 -credential $cred

.EXAMPLE
    # this will obtain from 192.168.100.1 with credential you enter.
    $cred = Get-Credential
    "server01","server02" | Get-AntiSpyware -credential $cred

.EXAMPLE
    # Output sample
    --------------------
    isplayName               : Windows Defender
    instanceGuid             : {D68DDC3A-831F-4fae-9E44-DA132C1ACF46}
    pathToSignedProductExe   : %ProgramFiles%\Windows Defender\MSASCui.exe
    pathToSignedReportingExe : %ProgramFiles%\Windows Defender\MsMpeng.exe
    productState             : 397568
    timestamp                : Fri, 25 Oct 2013 14:31:11 GMT
    PSComputerName           : 127.0.0.1
    --------------------    

#>

    [CmdletBinding()]
    Param
    (
        # Input ComputerName you want to check
        [Parameter(Mandatory = 0, 
                   ValueFromPipeline,
                   ValueFromPipelineByPropertyName, 
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $computerName = [System.Environment]::MachineName,

        # Input PSCredential for $ComputerName
        [Parameter(Mandatory = 0, 
                   Position=1)]
        [System.Management.Automation.PSCredential]
        $credential
    )

    Begin
    {
        $nameSpace = "SecurityCenter2"
        $className = "AntiSpywareProduct"
    }

    Process
    {
        if ($PSBoundParameters.count -eq 0)
        {
            if ((Get-CimInstance -namespace "root" -className "__Namespace").Name -contains $nameSpace)
            {
                Write-Verbose ("localhost cim session")
                Get-CimInstance -Namespace "root\$nameSpace" -ClassName $className
            }
            else
            {
                Write-Warning ("You can not check AntiSpyware with {0} as it not contain SecutiryCenter2" -f $OSName)
            }
        }
        else
        {
            try
            {
                Write-Verbose ("creating cim session for {0}" -f $computerName)
                $cimSession = New-CimSession @PSBoundParameters
                if ((Get-CimInstance -namespace "root" -className "__Namespace" -cimsession $cimSession).Name -contains $nameSpace)
                {
                    Get-CimInstance -Namespace "root\$nameSpace" -ClassName $className -CimSession $cimSession
                }
                else
                {
                    Write-Warning ("{0} not contains namespace {1}, you can not check {2}." -f $computerName, $nameSpace, $className)
                }
            }
            finally
            {
                $cimSession.Dispose()
            }
        }
    }

    End
    {
    }
}
## End Get-AntiSpyware
## Begin Get-AsciiReaction
Function Get-AsciiReaction
{
<#

	.SYNOPSIS
	
	Displays Ascii for different reactions and copies it to clipboard.

	.DESCRIPTION
	
	Displays Ascii for different reactions and copies it to clipboard.

	.EXAMPLE
	
	Get-AsciiReaction -Name Shrug
	
	Displays a shurg and copies it to clipboard.

	.NOTES
	
	Based on Reddit Thread https://www.reddit.com/r/PowerShell/comments/4aipw5/%E3%83%84/
	and Matt Hodge Function: https://github.com/MattHodge/MattHodgePowerShell/blob/master/Fun/Get-Ascii.ps1
#>
	[cmdletbinding()]
	Param
	(
		# Name of the Ascii 
		[Parameter()]
		[ValidateSet(
					 'Shrug',
					 'Disapproval',
					 'TableFlip',
					 'TableBack',
					 'TableFlip2',
					 'TableBack2',
					 'TableFlip3',
					 'Denko',
					 'BlowKiss',
					 'Lenny',
					 'Angry',
					 'DontKnow')]
		[string]$Name
	)
	
	$OutputEncoding = [System.Text.Encoding]::unicode
	
	# Function to write ascii to screen as well as clipboard it
	Function Write-Ascii
	{
		[CmdletBinding()]
		Param
		(
			# Ascii Data
			[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
			[string]$Ascii
		)
		
		# Clips it without the newline
		Add-Type -Assembly PresentationCore
		$clipText = ($Ascii).ToString() | Out-String -Stream
		[Windows.Clipboard]::SetText($clipText)
		
		Write-Output $clipText
	}
	
	Switch ($Name)
	{
		'Shrug' { [char[]]@(175, 92, 95, 40, 12484, 41, 95, 47, 175) -join '' | Write-Ascii }
		'Disapproval' { [char[]]@(3232, 95, 3232) -join '' | Write-Ascii }
		'TableFlip' { [char[]]@(40, 9583, 176, 9633, 176, 65289, 9583, 65077, 32, 9531, 9473, 9531, 41) -join '' | Write-Ascii }
		'TableBack' { [char[]]@(9516, 9472, 9472, 9516, 32, 175, 92, 95, 40, 12484, 41) -join '' | Write-Ascii }
		'TableFlip2' { [char[]]@(9531, 9473, 9531, 32, 65077, 12541, 40, 96, 1044, 180, 41, 65417, 65077, 32, 9531, 9473, 9531) -join '' | Write-Ascii }
		'TableBack2' { [char[]]@(9516, 9472, 9516, 12494, 40, 32, 186, 32, 95, 32, 186, 12494, 41) -join '' | Write-Ascii }
		'TableFlip3' { [char[]]@(40, 12494, 3232, 30410, 3232, 41, 12494, 24417, 9531, 9473, 9531) -join '' | Write-Ascii }
		'Denko' { [char[]]@(40, 180, 65381, 969, 65381, 96, 41) -join '' | Write-Ascii }
		'BlowKiss' { [char[]]@(40, 42, 94, 51, 94, 41, 47, 126, 9734) -join '' | Write-Ascii }
		'Lenny' { [char[]]@(40, 32, 865, 176, 32, 860, 662, 32, 865, 176, 41) -join '' | Write-Ascii }
		'Angry' { [char[]]@(40, 65283, 65439, 1044, 65439, 41) -join '' | Write-Ascii }
		'DontKnow' { [char[]]@(9488, 40, 39, 65374, 39, 65307, 41, 9484) -join '' | Write-Ascii }
		default
		{
			[PSCustomObject][ordered]@{
				'Shrug' = [char[]]@(175, 92, 95, 40, 12484, 41, 95, 47, 175) -join '' | Write-Ascii
				'Disapproval' = [char[]]@(3232, 95, 3232) -join '' | Write-Ascii
				'TableFlip' = [char[]]@(40, 9583, 176, 9633, 176, 65289, 9583, 65077, 32, 9531, 9473, 9531, 41) -join '' | Write-Ascii
				'TableBack' = [char[]]@(9516, 9472, 9472, 9516, 32, 175, 92, 95, 40, 12484, 41) -join '' | Write-Ascii 
				'TableFlip2' = [char[]]@(9531, 9473, 9531, 32, 65077, 12541, 40, 96, 1044, 180, 41, 65417, 65077, 32, 9531, 9473, 9531) -join '' | Write-Ascii 
				'TableBack2' = [char[]]@(9516, 9472, 9516, 12494, 40, 32, 186, 32, 95, 32, 186, 12494, 41) -join '' | Write-Ascii 
				'TableFlip3' = [char[]]@(40, 12494, 3232, 30410, 3232, 41, 12494, 24417, 9531, 9473, 9531) -join '' | Write-Ascii 
				'Denko' = [char[]]@(40, 180, 65381, 969, 65381, 96, 41) -join '' | Write-Ascii 
				'BlowKiss' = [char[]]@(40, 42, 94, 51, 94, 41, 47, 126, 9734) -join '' | Write-Ascii 
				'Lenny' = [char[]]@(40, 32, 865, 176, 32, 860, 662, 32, 865, 176, 41) -join '' | Write-Ascii 
				'Angry' = [char[]]@(40, 65283, 65439, 1044, 65439, 41) -join '' | Write-Ascii 
				'DontKnow' = [char[]]@(9488, 40, 39, 65374, 39, 65307, 41, 9484) -join '' | Write-Ascii 
			}
		}
	}
}
## End Get-AsciiReaction
## Begin Get-AntiVirusProduct
Function Get-AntiVirusProduct { 
[CmdletBinding()] 
param ( 
[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] 
[Alias('name')] 
$computername=$env:computername 
)
$AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct  -ComputerName $computername

#Switch to determine the status of antivirus definitions and real-time protection. 
#The values in this switch-statement are retrieved from the following website: http://community.kaseya.com/resources/m/knowexch/1020.aspx 
switch ($AntiVirusProduct.productState) { 
"262144" {$defstatus = "Up to date" ;$rtstatus = "Disabled"} 
    "262160" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "266240" {$defstatus = "Up to date" ;$rtstatus = "Enabled"} 
    "266256" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
    "393216" {$defstatus = "Up to date" ;$rtstatus = "Disabled"} 
    "393232" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "393488" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "397312" {$defstatus = "Up to date" ;$rtstatus = "Enabled"} 
    "397328" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
    "397584" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
default {$defstatus = "Unknown" ;$rtstatus = "Unknown"} 
    }

#Create hash-table for each computer 
$ht = @{} 
$ht.Computername = $computername 
$ht.Name = $AntiVirusProduct.displayName 
$ht.ProductExecutable = $AntiVirusProduct.pathToSignedProductExe 
$ht.'Definition Status' = $defstatus 
$ht.'Real-time Protection Status' = $rtstatus

#Create a new object for each computer 
New-Object -TypeName PSObject -Property $ht

}
## End Get-AntiVirusProduct
## Begin Get-AVStatus2
Function Get-AVStatus2 {

<#
.Synopsis
Get anti-virus product information
.Description
This command uses WMI via the Get-CimInstance command to query the state of installed anti-virus products. The default behavior is to only display enabled products, unless you use -All. You can query by computername or existing CIMSessions.
.Example
PS C:\> Get-AVStatus chi-win10

Displayname  : ESET NOD32 Antivirus 9.0.386.0
ProductState : 266256
Enabled      : True
UpToDate     : True
Path         : C:\Program Files\ESET\ESET NOD32 Antivirus\ecmd.exe
Timestamp    : Thu, 21 Jul 2016 15:20:18 GMT
Computername : CHI-WIN10

.Example
PS C:\>  import-csv s:\computers.csv | Get-AVStatus -All | Group Displayname | Select Name,Count | Sort Count,Name

Name                           Count
----                           -----
ESET NOD32 Antivirus 9.0.386.0    12
ESET Endpoint Security 5.0         6
Windows Defender                   4
360 Total Security                 1

Import a CSV file which includes a Computername heading. The imported objects are piped to this command. The results are sent to Group-Object.

.Example
PS C:\> $cs | Get-AVStatus | where {-Not $_.UptoDate}

Displayname  : ESET NOD32 Antivirus 9.0.386.0
ProductState : 266256
Enabled      : True
UpToDate     : False
Path         : C:\Program Files\ESET\ESET NOD32 Antivirus\ecmd.exe
Timestamp    : Wed, 20 Jul 2016 11:10:13 GMT
Computername : CHI-WIN11

Displayname  : ESET NOD32 Antivirus 9.0.386.0
ProductState : 266256
Enabled      : True
UpToDate     : False
Path         : C:\Program Files\ESET\ESET NOD32 Antivirus\ecmd.exe
Timestamp    : Thu, 07 Jul 2016 15:15:26 GMT
Computername : CHI-WIN81

You can also pipe CIMSession objects. In this example, the output are enabled products that are not up to date.
.Notes
version: 1.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

.Inputs
[string[]]
[Microsoft.Management.Infrastructure.CimSession[]]

.Outputs
[pscustomboject]

.Link
Get-CimInstance
#>

[cmdletbinding(DefaultParameterSetName="computer")]

Param(
[Parameter(
 Position = 0, 
 ValueFromPipeline, 
 ValueFromPipelineByPropertyName,
 ParameterSetName="computer")]
[ValidateNotNullorEmpty()]
#The name of a computer to query.
[string[]]$Computername = $env:COMPUTERNAME,

[Parameter(ValueFromPipeline,ParameterSetName = "session")]
#An existing CIMsession.
[Microsoft.Management.Infrastructure.CimSession[]]$WimSession,

#The default is enabled products only.
[switch]$All

)

Begin {
    Write-Verbose "[BEGIN  ] Starting: $($MyInvocation.Mycommand)"  

    Function ConvertTo-Hex {
    Param([int]$Number)
    '0x{0:x}' -f $Number
    }

    #initialize an hashtable of paramters to splat to Get-CimInstance
    $wimParams = @{
    Namespace = "root/SecurityCenter2"
    ClassName = "AntiVirusProduct"
    ErrorAction = "Stop"

    }

    If ($All) {
        Write-Verbose "[BEGIN  ] Getting all AV products"
    }
    
    $results = @()
} #begin

Process {
 
    #initialize an empty array to hold results
    $AV=@()
 
    #display PSBoundparameters formatted nicely for Verbose output  
    [string]$pb = ($PSBoundParameters | Format-Table -AutoSize | Out-String).TrimEnd()
    Write-Verbose "[PROCESS] PSBoundparameters: `n$($pb.split("`n").Foreach({"$("`t"*4)$_"}) | Out-String) `n" 
    Write-Verbose "[PROCESS] Using parameter set: $($pscmdlet.ParameterSetName)"

    if ($pscmdlet.ParameterSetName -eq 'computer') {
        foreach ($computer in $Computername) {

            Write-Verbose "[PROCESS] Querying $($computer.ToUpper())"
            $wimParams.ComputerName = $computer
            Try {    
                $AV += Get-WMIObject @wimParams
         
            }
            Catch {
                Write-Warning "[$($computer.ToUpper())] $($_.Exception.Message)"
                $wimParams.ComputerName = $null
            }

        } #foreach computer
    } 
    else {
        foreach ($session in $WimSession) {

            Write-Verbose "[PROCESS] Using session $($session.computername.toUpper())"
            $wimParams.CimSession = $session
            Try {    
                $AV += Get-WMIObject @wimParams
         
            }
            Catch {
                Write-Warning "[$($session.computername.ToUpper())] $($_.Exception.Message)"
                $wimParams.cimsession = $null
            }

        } #foreach computer
    }

       foreach ($item in $AV) {
                Write-Verbose "[PROCESS] Found $($item.Displayname)"
                $hx = ConvertTo-Hex $item.ProductState
                $mid = $hx.Substring(3,2)
                if ($mid -match "00|01") {
                    $Enabled = $False
                }
                else {
                    $Enabled = $True
                }
                $end = $hx.Substring(5)
                if ($end -eq "00") {
                    $UpToDate = $True
                }
                else {
                    $UpToDate = $False
                }

                $results += $item | Select Displayname,ProductState,
                @{Name="Enabled";Expression = {$Enabled}},
                @{Name = "UpToDate";Expression = {$UptoDate}},
                @{Name = "Path"; Expression = {$_.pathToSignedProductExe}},
                Timestamp,
                @{Name = "Computername"; Expression = {$_.PSComputername.toUpper()}}

            } #foreach

} #process

End {
    If ($All) {
      $results
    }
    else {
        #filter for enabled only
        ($results).Where({$_.enabled})
    }

    Write-Verbose "[END    ] Ending: $($MyInvocation.Mycommand)"
} #end

} 
## End Get-AVStatus2
## Begin Get-ByOwner
Function Get-ByOwner
{
    Get-ChildItem -Recurse C:\ | Get-ACL | Where {$_.Owner -Match $args[0] }
}
Function Get-CDROMDetails {                        
[cmdletbinding()]                        
param(                        
    [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]                        
    [string[]]$ComputerName = $env:COMPUTERNAME                        
)                        
            
begin {}                        
process {                        
    foreach($Computer in $COmputerName) {                        
    $object = New-Object –TypeName PSObject –Prop(@{                        
                'ComputerName'=$Computer.ToUpper();                        
                'CDROMDrive'= $null;                        
                'Manufacturer'=$null                        
               })                        
    if(!(Test-Connection -ComputerName $Computer -Count 1 -ErrorAction 0)) {                        
        Write-Verbose "$Computer is OFFLINE"                        
    }                        
    try {                        
        $cd = Get-WMIObject -Class Win32_CDROMDrive -ComputerName $Computer -ErrorAction Stop                        
    } catch {                        
        Write-Verbose "Failed to Query WMI Class"                        
        Continue;                        
    }                        
            
    $Object.CDROMDrive = $cd.Drive                        
    $Object.Manufacturer = $cd.caption                        
    $Object                           
            
    }                        
}                        
            
end {}               
}
## End Get-ByOwner
## Begin Get-ComputerHardwareSpecification
Function Get-ComputerHardwareSpecification
<#
.SYNOPSIS
Get the hardware specifications of a Windows computer.

.DESCRIPTION
Get the hardware specifications of a Windows computer including CPU, memory, and storage.
The Get-ComputerHardwareSpecification Function uses CIM to retrieve the following specific
information from a local or remote Windows computer.
CPU Model
Current CPU clock speed
Max CPU clock speed
Number of CPU sockets
Number of CPU cores
Number of logical processors
CPU hyperthreading
Total amount of physical RAM
Total amount of storage

.PARAMETER ComputerName
Enter a computer name

.PARAMETER Credential
Enter a credential to be used when connecting to the computer.

.EXAMPLE
Get-ComputerHardwareSpecification

ComputerName      : workstation01
CpuName           : Intel(R) Core(TM) i7-2600 CPU @ 3.40GHz
CurrentClockSpeed : 3401
MaxClockSpeed     : 3401
NumberofSockets   : 1
NumberofCores     : 4
LogicalProcessors : 8
HyperThreading    : True
Memory(GB)        : 16
Storage(GB)       : 697.96

.EXAMPLE
Get-ComputerHardwareSpecification -ComputerName server02

ComputerName      : server02
CpuName           : Intel(R) Core(TM) i7-2600 CPU @ 3.40GHz
CurrentClockSpeed : 3401
MaxClockSpeed     : 3401
NumberofSockets   : 1
NumberofCores     : 4
LogicalProcessors : 8
HyperThreading    : True
Memory(GB)        : 16
Storage(GB)       : 697.96

.NOTES
Created by: Jason Wasser @wasserja
Modified: 6/14/2017 02:18:45 PM 
Requires the New-ResilientCimSession Function
.LINK
New-ResilientCimSession 
https://gallery.technet.microsoft.com/scriptcenter/Establish-CimSession-in-b2166b02
.LINK
https://gallery.technet.microsoft.com/scriptcenter/Get-ComputerHardwareSpecifi-cf7df13d
#>
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )
    begin {}

    process {
        foreach ($Computer in $ComputerName) {
            $ErrorActionPreference = 'Stop'
            # Establishing CIM Session
            try {
                Write-Verbose -Message "Attempting to get the hardware specifications of $Computer"
                $CimSession = New-ResilientCimSession -ComputerName $Computer -Credential $Credential
                
                Write-Verbose -Message "Gathering CPU information of $Computer"
                $CPU = Get-CimInstance -ClassName win32_processor -CimSession $CimSession

                Write-Verbose -Message "Gathering memory information of $Computer"
                $Memory = Get-CimInstance -ClassName win32_operatingsystem -CimSession $CimSession
            
                Write-Verbose -Message "Gathering storage information of $Computer"
                $Disks = Get-CimInstance -ClassName win32_logicaldisk -Filter "DriveType = 3" -CimSession $CimSession
                $Storage = "{0:N2}" -f (($Disks | Measure-Object -Property Size -Sum).Sum / 1Gb) -as [decimal]
            
                # Building object properties
                $SystemProperties = [ordered]@{
                    ComputerName      = $Memory.PSComputerName
                    CpuName           = ($CPU | Select-Object -Property Name -First 1).Name
                    CurrentClockSpeed = ($CPU | Select-Object -Property CurrentClockSpeed -First 1).CurrentClockSpeed
                    MaxClockSpeed     = ($CPU | Select-Object -Property MaxClockSpeed -First 1).MaxClockSpeed
                    NumberofSockets   = $CPU.SocketDesignation.Count
                    NumberofCores     = ($CPU | Measure-Object -Property NumberofCores -Sum).Sum 
                    LogicalProcessors = ($CPU | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
                    HyperThreading    = ($CPU | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum -gt ($CPU | Measure-Object -Property NumberofCores -Sum).Sum 
                    'Memory(GB)'      = [int]($Memory.TotalVisibleMemorySize / 1Mb)
                    'Storage(GB)'     = $Storage
                }
                    
                $ComputerSpecs = New-Object -TypeName psobject -Property $SystemProperties
                $ComputerSpecs
                Remove-CimSession -CimSession $CimSession
            }
            catch {
                $ErrorActionPreference = 'Continue'
                Write-Error -Message "Unable to connect to $Computer"
            }
        }
    }
    end {}
}
## End Get-ComputerHardwareSpecification
## Begin Get-ComputerInfo
Function Get-ComputerInfo
{

<#
.SYNOPSIS
   This Function query some basic Operating System and Hardware Information from
   a local or remote machine.

.DESCRIPTION
   This Function query some basic Operating System and Hardware Information from
   a local or remote machine.
   It requires PowerShell version 3 for the Ordered Hashtable.

   The properties returned are the Computer Name (ComputerName),the Operating 
   System Name (OSName), Operating System Version (OSVersion), Memory Installed 
   on the Computer in GigaBytes (MemoryGB), the Number of 
   Processor(s) (NumberOfProcessors), Number of Socket(s) (NumberOfSockets),
   and Number of Core(s) (NumberOfCores).

   This Function as been tested against Windows Server 2000, 2003, 2008 and 2012

.PARAMETER ComputerName
   Specify a ComputerName or IP Address. Default is Localhost.

.PARAMETER ErrorLog
   Specify the full path of the Error log file. Default is .\Errors.log.

.PARAMETER Credential
   Specify the alternative credential to use

.EXAMPLE
   Get-ComputerInfo

   ComputerName       : XAVIER
   OSName             : Microsoft Windows 8 Pro
   OSVersion          : 6.2.9200
   MemoryGB           : 4
   NumberOfProcessors : 1
   NumberOfSockets    : 1
   NumberOfCores      : 4

   This example return information about the localhost. By Default, if you don't
   specify a ComputerName, the Function will run against the localhost.

.EXAMPLE
   Get-ComputerInfo -ComputerName SERVER01

   ComputerName       : SERVER01
   OSName             : Microsoft Windows Server 2012
   OSVersion          : 6.2.9200
   MemoryGB           : 4
   NumberOfProcessors : 1
   NumberOfSockets    : 1
   NumberOfCores      : 4

   This example return information about the remote computer SERVER01.

.EXAMPLE
   Get-Content c:\ServersList.txt | Get-ComputerInfo
    
   ComputerName       : DC
   OSName             : Microsoft Windows Server 2012
   OSVersion          : 6.2.9200
   MemoryGB           : 8
   NumberOfProcessors : 1
   NumberOfSockets    : 1
   NumberOfCores      : 4

   ComputerName       : FILESERVER
   OSName             : Microsoft Windows Server 2008 R2 Standard 
   OSVersion          : 6.1.7601
   MemoryGB           : 2
   NumberOfProcessors : 1
   NumberOfSockets    : 1
   NumberOfCores      : 1

   ComputerName       : SHAREPOINT
   OSName             : Microsoft(R) Windows(R) Server 2003 Standard x64 Edition
   OSVersion          : 5.2.3790
   MemoryGB           : 8
   NumberOfProcessors : 8
   NumberOfSockets    : 8
   NumberOfCores      : 8

   ComputerName       : FTP
   OSName             : Microsoft Windows 2000 Server
   OSVersion          : 5.0.2195
   MemoryGB           : 4
   NumberOfProcessors : 2
   NumberOfSockets    : 2
   NumberOfCores      : 2

   This example show how to use the Function Get-ComputerInfo in a Pipeline.
   Get-Content Cmdlet Gather the content of the ServersList.txt and send the
   output to Get-ComputerInfo via the Pipeline.

.EXAMPLE
   Get-ComputerInfo -ComputerName FILESERVER,SHAREPOINT -ErrorLog d:\MyErrors.log.

   ComputerName       : FILESERVER
   OSName             : Microsoft Windows Server 2008 R2 Standard 
   OSVersion          : 6.1.7601
   MemoryGB           : 2
   NumberOfProcessors : 1
   NumberOfSockets    : 1
   NumberOfCores      : 1

   ComputerName       : SHAREPOINT
   OSName             : Microsoft(R) Windows(R) Server 2003 Standard x64 Edition
   OSVersion          : 5.2.3790
   MemoryGB           : 8
   NumberOfProcessors : 8
   NumberOfSockets    : 8
   NumberOfCores      : 8

   This example show how to use the Function Get-ComputerInfo against multiple
   Computers. Using the ErrorLog Parameter, we send the potential errors in the
   file d:\Myerrors.log.

.INPUTS
   System.String

.OUTPUTS
   System.Management.Automation.PSCustomObject

.NOTES
   Scripting Games 2013 - Advanced Event #2
#>

 [CmdletBinding()]

    PARAM(
    [Parameter(ValueFromPipeline=$true)]
    [String[]]$ComputerName = "LocalHost",

    [String]$ErrorLog = ".\Errors.log",

    [Alias("RunAs")]
    [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
    )#PARAM

    BEGIN {}#PROCESS BEGIN

    PROCESS{
        FOREACH ($Computer in $ComputerName) {
            Write-Verbose -Message "PROCESS - Querying $Computer ..."
                
            TRY{
                $Splatting = @{
                    ComputerName = $Computer
                }

                IF ($PSBoundParameters["Credential"]){
                    $Splatting.Credential = $Credential
                }


                $Everything_is_OK = $true
                Write-Verbose -Message "PROCESS - $Computer - Testing Connection"
                Test-Connection -Count 1 -ComputerName $Computer -ErrorAction Stop -ErrorVariable ProcessError | Out-Null

                # Query WMI class Win32_OperatingSystem
                Write-Verbose -Message "PROCESS - $Computer - WMI:Win32_OperatingSystem"
                $OperatingSystem = Get-WmiObject -Class Win32_OperatingSystem @Splatting -ErrorAction Stop -ErrorVariable ProcessError
                    
                # Query WMI class Win32_ComputerSystem
                Write-Verbose -Message "PROCESS - $Computer - WMI:Win32_ComputerSystem"
                $ComputerSystem = Get-WmiObject -Class win32_ComputerSystem @Splatting -ErrorAction Stop -ErrorVariable ProcessError

                # Query WMI class Win32_Processor
                Write-Verbose -Message "PROCESS - $Computer - WMI:Win32_Processor"
                $Processors = Get-WmiObject -Class win32_Processor @Splatting -ErrorAction Stop -ErrorVariable ProcessError

                # Processors - Determine the number of Socket(s) and core(s)
                # The following code is required for some old Operating System where the
                # property NumberOfCores does not exist.
                Write-Verbose -Message "PROCESS - $Computer - Determine the number of Socket(s)/Core(s)"
                $Cores = 0
                $Sockets = 0
                FOREACH ($Proc in $Processors){
                    IF($Proc.numberofcores -eq $null){
                        IF ($Proc.SocketDesignation -ne $null){$Sockets++}
                        $Cores++
                    }ELSE {
                        $Sockets++
                        $Cores += $proc.numberofcores
                    }#ELSE
                }#FOREACH $Proc in $Processors

            }CATCH{
                $Everything_is_OK = $false
                Write-Warning -Message "Error on $Computer"
                $Computer | Out-file -FilePath $ErrorLog -Append -ErrorAction Continue
                $ProcessError | Out-file -FilePath $ErrorLog -Append -ErrorAction Continue
                Write-Warning -Message "Logged in $ErrorLog"

            }#CATCH


            IF ($Everything_is_OK){
                Write-Verbose -Message "PROCESS - $Computer - Building the Output Information"
                $Info = [ordered]@{
                    "ComputerName" = $OperatingSystem.__Server;
                    "OSName" = $OperatingSystem.Caption;
                    "OSVersion" = $OperatingSystem.version;
                    "MemoryGB" = $ComputerSystem.TotalPhysicalMemory/1GB -as [int];
                    "NumberOfProcessors" = $ComputerSystem.NumberOfProcessors;
                    "NumberOfSockets" = $Sockets;
                    "NumberOfCores" = $Cores}

                $output = New-Object -TypeName PSObject -Property $Info
                $output
            } #end IF Everything_is_OK
        }#end Foreach $Computer in $ComputerName
    }#PROCESS BLOCK
    END{
        # Cleanup
        Write-Verbose -Message "END - Cleanup Variables"
        Remove-Variable -Name output,info,ProcessError,Sockets,Cores,OperatingSystem,ComputerSystem,Processors,
        ComputerName, ComputerName, Computer, Everything_is_OK -ErrorAction SilentlyContinue
        
        # End
        Write-Verbose -Message "END - Script End !"
    }#END BLOCK
}
## End Get-ComputerInfo
## Begin Get-ComputerOS
Function Get-ComputerOS
{
<#
	.SYNOPSIS
		Function to retrieve the Operating System of a machine
	
	.DESCRIPTION
		Function to retrieve the Operating System of a machine
	
	.PARAMETER ComputerName
		Specifies the ComputerName of the machine to query. Default is localhost.
	
	.PARAMETER Credential
		Specifies the credentials to use. Default is Current credentials
	
	.EXAMPLE
		PS C:\> Get-ComputerOS -ComputerName "SERVER01","SERVER02","SERVER03"
	
	.EXAMPLE
		PS C:\> Get-ComputerOS -ComputerName "SERVER01" -Credential (Get-Credential -cred "FX\SuperAdmin")
	
	.NOTES
		Additional information about the Function.
#>
	[CmdletBinding()]
	PARAM (
		[Parameter(ParameterSetName = "Main")]
		[Alias("CN","__SERVER","PSComputerName")]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Parameter(ParameterSetName="Main")]
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty,
		
		[Parameter(ParameterSetName = "CimSession")]
		[Microsoft.Management.Infrastructure.CimSession]$CimSession
	)
	BEGIN
	{
		# Default Verbose/Debug message
		Function Get-DefaultMessage
		{
	<#
	.SYNOPSIS
		Helper Function to show default message used in VERBOSE/DEBUG/WARNING
	.DESCRIPTION
		Helper Function to show default message used in VERBOSE/DEBUG/WARNING.
		Typically called inside another Function in the BEGIN Block
	#>
			PARAM ($Message)
			Write-Output "[$(Get-Date -Format 'yyyy/MM/dd-HH:mm:ss:ff')][$((Get-Variable -Scope 1 -Name MyInvocation -ValueOnly).MyCommand.Name)] $Message"
		}#Get-DefaultMessage
	}
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				Write-Verbose -Message (Get-DefaultMessage -Message $Computer)
				IF (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
				{
					# Define Hashtable to hold our properties
					$Splatting = @{
						class = "Win32_OperatingSystem"
						ErrorAction = Stop
					}
					
					IF ($PSBoundParameters['CimSession'])
					{
						Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - CimSession")
						# Using cim session already opened
						$Query = Get-CIMInstance @Splatting -CimSession $CimSession
					}
					ELSE
					{
						# Credential specified
						IF ($PSBoundParameters['Credential'])
						{
							Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Credential specified $($Credential.username)")
							$Splatting.Credential = $Credential
						}
						
						# Set the ComputerName into the splatting
						$Splatting.ComputerName = $ComputerName
						Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Get-WmiObject")
						$Query = Get-WmiObject @Splatting
					}
					
					# Prepare output
					$Properties = @{
						ComputerName = $Computer
						OperatingSystem = $Query.Caption
					}
					
					# Output
					New-Object -TypeName PSObject -Property $Properties
				}
			}
			CATCH
			{
				Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Issue to connect")
				Write-Verbose -Message $Error[0].Exception.Message
			}#CATCH
			FINALLY
			{
				$Splatting.Clear()
			}
		}#FOREACH
	}#PROCESS
	END
	{
		Write-Warning -Message (Get-DefaultMessage -Message "Script completed")
	}
}
## End Get-ComputerOS
## Begin Get-ComputerVirtualStatus
Function Get-ComputerVirtualStatus {
    <# 
    .SYNOPSIS 
    Validate if a remote server is virtual or physical 
    .DESCRIPTION 
    Uses wmi (along with an optional credential) to determine if a remote computers, or list of remote computers are virtual. 
    If found to be virtual, a best guess effort is done on which type of virtual platform it is running on. 
    .PARAMETER ComputerName 
    Computer or IP address of machine 
    .PARAMETER Credential 
    Provide an alternate credential 
    .EXAMPLE 
    $Credential = Get-Credential 
    Get-RemoteServerVirtualStatus 'Server1','Server2' -Credential $Credential | select ComputerName,IsVirtual,VirtualType | ft 
     
    Description: 
    ------------------ 
    Using an alternate credential, determine if server1 and server2 are virtual. Return the results along with the type of virtual machine it might be. 
    .EXAMPLE 
    (Get-RemoteServerVirtualStatus server1).IsVirtual 
     
    Description: 
    ------------------ 
    Determine if server1 is virtual and returns either true or false. 

    .LINK 
    http://www.the-little-things.net/ 
    .LINK 
    http://nl.linkedin.com/in/zloeber 
    .NOTES 
     
    Name       : Get-RemoteServerVirtualStatus 
    Version    : 1.1.0 12/09/2014
                 - Removed prompt for credential
                 - Refactored some of the code a bit.
                 1.0.0 07/27/2013 
                 - First release 
    Author     : Zachary Loeber 
    #> 
    [cmdletBinding(SupportsShouldProcess = $true)] 
    param( 
        [parameter(Position=0, ValueFromPipeline=$true, HelpMessage="Computer or IP address of machine to test")] 
        [string[]]$ComputerName = $env:COMPUTERNAME, 
        [parameter(HelpMessage="Pass an alternate credential")] 
        [System.Management.Automation.PSCredential]$Credential = $null 
    ) 
    begin {
        $WMISplat = @{} 
        if ($Credential -ne $null) { 
            $WMISplat.Credential = $Credential 
        } 
        $results = @()
        $computernames = @()
    } 
    process { 
        $computernames += $ComputerName 
    } 
    end {
        foreach($computer in $computernames) { 
            $WMISplat.ComputerName = $computer 
            try { 
                $wmibios = Get-WmiObject Win32_BIOS @WMISplat -ErrorAction Stop | Select-Object version,serialnumber 
                $wmisystem = Get-WmiObject Win32_ComputerSystem @WMISplat -ErrorAction Stop | Select-Object model,manufacturer
                $ResultProps = @{
                    ComputerName = $computer 
                    BIOSVersion = $wmibios.Version 
                    SerialNumber = $wmibios.serialnumber 
                    Manufacturer = $wmisystem.manufacturer 
                    Model = $wmisystem.model 
                    IsVirtual = $false 
                    VirtualType = $null 
                }
                if ($wmibios.SerialNumber -like "*VMware*") {
                    $ResultProps.IsVirtual = $true
                    $ResultProps.VirtualType = "Virtual - VMWare"
                }
                else {
                    switch -wildcard ($wmibios.Version) {
                        'VIRTUAL' { 
                            $ResultProps.IsVirtual = $true 
                            $ResultProps.VirtualType = "Virtual - Hyper-V" 
                        } 
                        'A M I' {
                            $ResultProps.IsVirtual = $true 
                            $ResultProps.VirtualType = "Virtual - Virtual PC" 
                        } 
                        '*Xen*' { 
                            $ResultProps.IsVirtual = $true 
                            $ResultProps.VirtualType = "Virtual - Xen" 
                        }
                    }
                }
                if (-not $ResultProps.IsVirtual) {
                    if ($wmisystem.manufacturer -like "*Microsoft*") 
                    { 
                        $ResultProps.IsVirtual = $true 
                        $ResultProps.VirtualType = "Virtual - Hyper-V" 
                    } 
                    elseif ($wmisystem.manufacturer -like "*VMWare*") 
                    { 
                        $ResultProps.IsVirtual = $true 
                        $ResultProps.VirtualType = "Virtual - VMWare" 
                    } 
                    elseif ($wmisystem.model -like "*Virtual*") { 
                        $ResultProps.IsVirtual = $true
                        $ResultProps.VirtualType = "Unknown Virtual Machine"
                    }
                }
                $results += New-Object PsObject -Property $ResultProps
            }
            catch {
                Write-Warning "Cannot connect to $computer"
            } 
        } 
        return $results 
    } 
}
## End Get-ComputerVirtualStatus
## Begin Get-DCDiag
Function Get-DCDiag {
          
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
             
}
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

## End Get-DCDiag
## Begin Get-DefragAnalysis
Function Get-DefragAnalysis {

<#
.Synopsis
Run a defrag analysis.
.Description
This command uses WMI to to run a defrag analysis on selected volumes on
local or remote computers. You will get a custom object for each volume like
this:

AverageFileSize               : 64
AverageFragmentsPerFile       : 1
AverageFreeSpacePerExtent     : 17002496
ClusterSize                   : 4096
ExcessFolderFragments         : 0
FilePercentFragmentation      : 0
FragmentedFolders             : 0
FreeSpace                     : 161816576
FreeSpacePercent              : 77
FreeSpacePercentFragmentation : 29
LargestFreeSpaceExtent        : 113500160
MFTPercentInUse               : 100
MFTRecordCount                : 511
PageFileSize                  : 0
TotalExcessFragments          : 0
TotalFiles                    : 182
TotalFolders                  : 11
TotalFragmentedFiles          : 0
TotalFreeSpaceExtents         : 8
TotalMFTFragments             : 1
TotalMFTSize                  : 524288
TotalPageFileFragments        : 0
TotalPercentFragmentation     : 0
TotalUnmovableFiles           : 4
UsedSpace                     : 47894528
VolumeName                    : 
VolumeSize                    : 209711104
Driveletter                   : E:
DefragRecommended             : False
Computername                  : NOVO8

The default drive is C: on the local computer.
.Example
PS C:\> Get-DefragAnalysis
Run a defrag analysis on C: on the local computer
.Example
PS C:\> Get-DefragAnalysis -drive "C:" -computername $servers
Run a defrag analysis for drive C: on a previously defined collection of server names.
.Example
PS C:\> $data = Get-WmiObject Win32_volume -filter "driveletter like '%' AND drivetype=3" -ComputerName Novo8 | Get-DefragAnalysis
PS C:\> $data | Sort Driveletter | Select Computername,DriveLetter,DefragRecommended

Computername                    Driveletter                     DefragRecommended
------------                    -----------                     -----------------
NOVO8                           C:                                          False
NOVO8                           D:                                          True
NOVO8                           E:                                          False

Get all volumes on a remote computer that are fixed but have a drive letter,
this should eliminate CD/DVD drives, and run a defrag analysis on each one.
The results are saved to a variable, $data.
.Notes
Last Updated: 12/5/2012
Author      : Jeffery Hicks (http://jdhitsolutions.com/blog)
Version     : 0.9

.Link
Get-WMIObject
Invoke-WMIMethod

#>

[cmdletbinding(SupportsShouldProcess=$True)]

Param(
[Parameter(Position=0,ValueFromPipelineByPropertyName=$True)]
[ValidateNotNullorEmpty()]
[Alias("drive")]
[string]$Driveletter="C:",
[Parameter(Position=1,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[ValidateNotNullorEmpty()]
[Alias("PSComputername","SystemName")]
[string[]]$Computername=$env:computername
)

Begin {
    Write-Verbose -Message "$(Get-Date) Starting $($MyInvocation.Mycommand)"   
} #close Begin

Process {
    #strip off any extra spaces on the drive letter just in case
    Write-Verbose "$(Get-Date) Processing $Driveletter"
    $Driveletter=$Driveletter.Trim()
    if ($Driveletter.length -gt 2) {
        Write-Verbose "$(Get-Date) Scrubbing drive parameter value"
        $Driveletter=$Driveletter.Substring(0,2)
    }
    #add a colon if not included
    if ($Driveletter -match "^\w$") {
        Write-Verbose "$(Get-Date) Modifying drive parameter value"
        $Driveletter="$($Driveletter):"
    }

    Write-Verbose "$(Get-Date) Analyzing drive $Driveletter"
        
    Foreach ($computer in $computername) {
        Write-Verbose "$(Get-Date) Examining $computer"
        Try {
            $volume=Get-WmiObject -Class Win32_Volume -filter "DriveLetter='$Driveletter'" -computername $computer -errorAction "Stop"
        }
        Catch {
            Write-Warning ("Failed to get volume {0} from  {1}. {2}" -f $driveletter,$computer,$_.Exception.Message)
        }
        if ($volume) {
            Write-Verbose "$(Get-Date) Running defrag analysis"
            $analysis = $volume | Invoke-WMIMethod -name DefragAnalysis
        
            #get properties for DefragAnalysis so we can filter out system properties
            $analysis.DefragAnalysis.Properties | 
            Foreach -begin {$Prop=@()} -process { $Prop+=$_.Name }
        
            Write-Verbose "$(Get-Date) Retrieving results"
            $analysis | Select @{Name="Results";Expression={$_.DefragAnalysis | 
            Select-Object -Property $Prop |
            Foreach-Object { 
              #Add on some additional property values
              $_ | Add-member -MemberType Noteproperty -Name Driveletter -value $DriveLetter
              $_ | Add-member -MemberType Noteproperty -Name DefragRecommended -value $analysis.DefragRecommended 
              $_ | Add-member -MemberType Noteproperty -Name Computername -value $volume.__SERVER -passthru
             } #foreach-object
            }}  | Select -expand Results 
            
            #clean up variables so there are no accidental leftovers
            Remove-Variable "volume","analysis"
        } #close if volume
     } #close Foreach computer
 } #close Process
 
End {
    Write-Verbose "$(Get-Date) Defrag analysis complete"
} #close End

}
## End Get-DefragAnalysis
## Begin Get-DirectoryVolume
Function Get-DirectoryVolume
{

    [CmdletBinding()]
    param
    (
        [parameter(
            position = 0,
            mandatory = 1,
            valuefrompipeline = 1,
            valuefrompipelinebypropertyname = 1)]
        [string[]]
        $Path = $null,

        [parameter(
            position = 1,
            mandatory = 0,
            valuefrompipelinebypropertyname = 1)]
        [validateSet("KB", "MB", "GB")]
        [string]
        $Scale = "KB",

        [parameter(
            position = 2,
            mandatory = 0,
            valuefrompipelinebypropertyname = 1)]
        [switch]
        $Recurse = $false,

        [parameter(
            position = 3,
            mandatory = 0,
            valuefrompipelinebypropertyname = 1)]
        [switch]
        $Ascending = $false,

        [parameter(
            position = 4,
            mandatory = 0,
            valuefrompipelinebypropertyname = 1)]
        [switch]
        $OmitZero
    )

    process
    {
        $path `
        | %{
            Write-Verbose ("Checking path : {0}. Scale : {1}. Recurse switch : {2}. Decending : {3}" -f $_, $Scale, $Recurse, !$Ascending)
            if (Test-Path $_)
            {
                $result = Get-ChildItem -Path $_ -Recurse:$Recurse `
                | where PSIsContainer `
                | %{
                    $subFolderItems = (Get-ChildItem $_.FullName | where Length | measure Length -sum)
                    [PSCustomObject]@{
                        Fullname = $_.FullName
                        $scale = [decimal]("{0:N4}" -f ($subFolderItems.sum / "1{0}" -f $scale))
                    }} `
                | sort $scale -Descending:(!$Ascending)

                if ($OmitZero)
                {
                    return $result | where $Scale -ne ([decimal]("{0:N4}" -f "0.0000"))
                }
                else
                {
                    return $result
                }
            }
        }
    }
}
## End Get-DirectoryVolume
## Begin Get-GitCurrentRelease
Function Get-GitCurrentRelease
<#
# Oneliner
Get-ChildItem c:\logs -Recurse | where PSIsContainer | %{$i=$_;$subFolderItems = (Get-ChildItem $i.FullName | where Length | measure Length -sum);[PSCustomObject]@{Fullname=$i.FullName;MB=[decimal]("{0:N2}" -f ($subFolderItems.sum / 1MB))}} | sort MB -Descending | ft -AutoSize

# refine oneliner
Get-ChildItem c:\logs -Recurse `
| where PSIsContainer `
| %{
    $i=$_
    $subFolderItems = (Get-ChildItem $i.FullName | where Length | measure Length -sum)
    [PSCustomObject]@{
        Fullname=$i.FullName
        MB=[decimal]("{0:N2}" -f ($subFolderItems.sum / 1MB))
    }} `
| sort MB -Descending `
| format -AutoSize


# if devide each
$folder = Get-ChildItem c:\logs -recurse | where PSIsContainer
[array]$volume = foreach ($i in $folder)
{
    $subFolderItems = (Get-ChildItem $i.FullName | where Length | measure Length -sum)
    [PSCustomObject]@{
        Fullname=$i.FullName
        MB=[decimal]("{0:N2}" -f ($subFolderItems.sum / 1MB))
    }
}
$volume | sort MB -Descending | ft -AutoSize
18:53

#>
{
[cmdletbinding()]
Param(
[ValidateNotNullorEmpty()]
[string]$Uri = "https://api.github.com/repos/git-for-windows/git/releases/latest"
)
 
Begin {
    Write-Verbose "[BEGIN  ] Starting: $($MyInvocation.Mycommand)"  
 
} #begin
 
Process {
    Write-Verbose "[PROCESS] Getting current release information from $uri"
    $data = Invoke-Restmethod -uri $uri -Method Get
 
    
    if ($data.tag_name) {
    [pscustomobject]@{
        Name = $data.name
        Version = $data.tag_name
        Released = $($data.published_at -as [datetime])
      }
   } 
} #process
 
End {
    Write-Verbose "[END    ] Ending: $($MyInvocation.Mycommand)"
} #end
 
}
## End Get-GitCurrentRelease
## Begin Get-HashTableEmptyValue
Function Get-HashTableEmptyValue
{
<#
.SYNOPSIS
    This Function will get the empty or Null entry of a hashtable object
.DESCRIPTION
    This Function will get the empty or Null entry of a hashtable object
.PARAMETER Hashtable
    Specifies the hashtable that will be showed
.EXAMPLE
    Get-HashTableEmptyValue -HashTable $SplattingVariable
.NOTES
    Francois-Xavier Cat
    @lazywinadm
    www.lazywinadmin.com
#>
    PARAM([System.Collections.Hashtable]$HashTable)

    $HashTable.GetEnumerator().name |
        ForEach-Object -Process {
            if($HashTable[$_] -eq "" -or $HashTable[$_] -eq $null)
            {
                Write-Output $_
            }
        }
}
## End Get-HashTableEmptyValue
## Begin Get-HashTableNotEmptyOrNullValue
Function Get-HashTableNotEmptyOrNullValue
{
<#
.SYNOPSIS
    This Function will get the values that are not empty or Null in a hashtable object
.DESCRIPTION
    This Function will get the values that are not empty or Null in a hashtable object
.PARAMETER Hashtable
    Specifies the hashtable that will be showed
.EXAMPLE
    Get-HashTableNotEmptyOrNullValue -HashTable $SplattingVariable
.NOTES
    Francois-Xavier Cat
    @lazywinadm
    www.lazywinadmin.com
#>
    PARAM([System.Collections.Hashtable]$HashTable)

    $HashTable.GetEnumerator().name |
        ForEach-Object -Process {
            if($HashTable[$_] -ne "")
            {
                Write-Output $_
            }
        }
}
## End Get-HashTableNotEmptyOrNullValue
Function Get-ImageInformation
{
<#
.SYNOPSIS
	Function to retrieve Image file information

.DESCRIPTION
	Function to retrieve Image file information

.PARAMETER FilePath
	Specify one or multiple image file path(s).

.EXAMPLE
	PS C:\> Get-ImageInformation -FilePath c:\temp\image.png

.NOTES
	Francois-Xavier Cat
	lazywinadmin.com
	@lazywinadm
	github.com/lazywinadmin
#>
	PARAM (
		[System.String[]]$FilePath
	)
	Foreach ($Image in $FilePath)
	{
		# Load Assembly
		Add-type -AssemblyName System.Drawing
		
		# Retrieve information
		New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Image
	}
}
<#
.SYNOPSIS
	Get-InstalledSoftware retrieves a list of installed software
.DESCRIPTION
	Get-InstalledSoftware opens up the specified (remote) registry and scours it for installed software. When found it returns a list of the software and it's version.
.PARAMETER ComputerName
	The computer from which you want to get a list of installed software. Defaults to the local host.
.EXAMPLE
	Get-InstalledSoftware DC1
	
	This will return a list of software from DC1. Like:
	Name			Version		Computer  UninstallCommand
	----			-------     --------  ----------------
	7-Zip 			9.20.00.0	DC1       MsiExec.exe /I{23170F69-40C1-2702-0920-000001000000}
	Google Chrome	65.119.95	DC1       MsiExec.exe /X{6B50D4E7-A873-3102-A1F9-CD5B17976208}
	Opera			12.16		DC1		  "C:\Program Files (x86)\Opera\Opera.exe" /uninstall
.EXAMPLE
	Import-Module ActiveDirectory
	Get-ADComputer -filter 'name -like "DC*"' | Get-InstalledSoftware
	
	This will get a list of installed software on every AD computer that matches the AD filter (So all computers with names starting with DC)
.INPUTS
	[string[]]Computername
.OUTPUTS
	PSObject with properties: Name,Version,Computer,UninstallCommand
.NOTES
	Author: Anthony Howell
	
	To add directories, add to the LMkeys (LocalMachine)
.LINK
	[Microsoft.Win32.RegistryHive]
	[Microsoft.Win32.RegistryKey]
#>
Function Get-InstalledSoftware
{
	Param
	(
		[Alias('Computer','ComputerName','HostName')]
		[Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
		[string[]]$Name = $env:COMPUTERNAME
	)
	Begin
	{
		$LMkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
		$LMtype = [Microsoft.Win32.RegistryHive]::LocalMachine
		$CUkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
		$CUtype = [Microsoft.Win32.RegistryHive]::CurrentUser
		
	}
	Process
	{
		ForEach($Computer in $Name)
		{
			$MasterKeys = @()
			If(!(Test-Connection -ComputerName $Computer -count 1 -quiet))
			{
				Write-Error -Message "Unable to contact $Computer. Please verify its network connectivity and try again." -Category ObjectNotFound -TargetObject $Computer
				Break
			}
			$CURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($CUtype,$computer)
			$LMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($LMtype,$computer)
			ForEach($Key in $LMkeys)
			{
				$RegKey = $LMRegKey.OpenSubkey($key)
				If($RegKey -ne $null)
				{
					ForEach($subName in $RegKey.getsubkeynames())
					{
						foreach($sub in $RegKey.opensubkey($subName))
						{
							$MasterKeys += (New-Object PSObject -Property @{
							"ComputerName" = $Computer
							"Name" = $sub.getvalue("displayname")
							"SystemComponent" = $sub.getvalue("systemcomponent")
							"ParentKeyName" = $sub.getvalue("parentkeyname")
							"Version" = $sub.getvalue("DisplayVersion")
							"UninstallCommand" = $sub.getvalue("UninstallString")
							})
						}
					}
				}
			}
			ForEach($Key in $CUKeys)
			{
				$RegKey = $CURegKey.OpenSubkey($Key)
				If($RegKey -ne $null)
				{
					ForEach($subName in $RegKey.getsubkeynames())
					{
						foreach($sub in $RegKey.opensubkey($subName))
						{
							$MasterKeys += (New-Object PSObject -Property @{
							"ComputerName" = $Computer
							"Name" = $sub.getvalue("displayname")
							"SystemComponent" = $sub.getvalue("systemcomponent")
							"ParentKeyName" = $sub.getvalue("parentkeyname")
							"Version" = $sub.getvalue("DisplayVersion")
							"UninstallCommand" = $sub.getvalue("UninstallString")
							})
						}
					}
				}
			}
			$MasterKeys = ($MasterKeys | Where {$_.Name -ne $Null -AND $_.SystemComponent -ne "1" -AND $_.ParentKeyName -eq $Null} | select Name,Version,ComputerName,UninstallCommand | sort Name)
			$MasterKeys
		}
	}
	End
	{
		
	}
}
Function Get-ISEShortCut
{
<#
.SYNOPSIS
	List ISE Shortcuts

.DESCRIPTION
    List ISE Shortcuts.
    This won't run in a regular powershell console, only in ISE.
	
.EXAMPLE
    Get-ISEShortcut

    Will list all the shortcuts available
.EXAMPLE
    Get-Help Get-ISEShortcut -Online

    Will show technet page of ISE Shortcuts
.LINK
	http://technet.microsoft.com/en-us/library/jj984298.aspx
	
.NOTES
	Francois-Xavier Cat
	www.lazywinadmin.com
	@lazywinadm
	
	VERSION HISTORY
	2015/01/10 Initial Version
#>
	PARAM($Key,$Name)
	BEGIN
	{
		Function Test-IsISE
		{
			# try...catch accounts for:
			# Set-StrictMode -Version latest
			try
			{
				return $psISE -ne $null;
			}
			catch
			{
				return $false;
			}
		}
	}
	PROCESS
	{
		if ($(Test-IsISE) -eq $true)
		{
			# http://www.powershellmagazine.com/2013/01/29/the-complete-list-of-powershell-ise-3-0-keyboard-shortcuts/
			
			# Reference to the ISE Microsoft.PowerShell.GPowerShell assembly (DLL)
			$gps = $psISE.GetType().Assembly
			$rm = New-Object System.Resources.ResourceManager GuiStrings, $gps
			$rs = $rm.GetResourceSet((Get-Culture), $true, $true)
			$rs | Where-Object Name -match 'Shortcut\d?$|^F\d+Keyboard' |
			Sort-Object Value
			
		}
	}
}
Function Get-LastLogon
{
<#

.SYNOPSIS
	This Function will list the last user logged on or logged in.

.DESCRIPTION
	This Function will list the last user logged on or logged in.  It will detect if the user is currently logged on
	via WMI or the Registry, depending on what version of Windows is running on the target.  There is some "guess" work
	to determine what Domain the user truly belongs to if run against Vista NON SP1 and below, since the Function
	is using the profile name initially to detect the user name.  It then compares the profile name and the Security
	Entries (ACE-SDDL) to see if they are equal to determine Domain and if the profile is loaded via the Registry.

.PARAMETER ComputerName
	A single Computer or an array of computer names.  The default is localhost ($env:COMPUTERNAME).

.PARAMETER FilterSID
	Filters a single SID from the results.  For use if there is a service account commonly used.
	
.PARAMETER WQLFilter
	Default WQLFilter defined for the Win32_UserProfile query, it is best to leave this alone, unless you know what
	you are doing.
	Default Value = "NOT SID = 'S-1-5-18' AND NOT SID = 'S-1-5-19' AND NOT SID = 'S-1-5-20'"
	
.EXAMPLE
	$Servers = Get-Content "C:\ServerList.txt"
	Get-LastLogon -ComputerName $Servers

	This example will return the last logon information from all the servers in the C:\ServerList.txt file.

	Computer          : SVR01
	User              : WILHITE\BRIAN
	SID               : S-1-5-21-012345678-0123456789-012345678-012345
	Time              : 9/20/2012 1:07:58 PM
	CurrentlyLoggedOn : False

	Computer          : SVR02
	User              : WILIHTE\BRIAN
	SID               : S-1-5-21-012345678-0123456789-012345678-012345
	Time              : 9/20/2012 12:46:48 PM
	CurrentlyLoggedOn : True
	
.EXAMPLE
	Get-LastLogon -ComputerName svr01, svr02 -FilterSID S-1-5-21-012345678-0123456789-012345678-012345

	This example will return the last logon information from all the servers in the C:\ServerList.txt file.

	Computer          : SVR01
	User              : WILHITE\ADMIN
	SID               : S-1-5-21-012345678-0123456789-012345678-543210
	Time              : 9/20/2012 1:07:58 PM
	CurrentlyLoggedOn : False

	Computer          : SVR02
	User              : WILIHTE\ADMIN
	SID               : S-1-5-21-012345678-0123456789-012345678-543210
	Time              : 9/20/2012 12:46:48 PM
	CurrentlyLoggedOn : True

.LINK
	http://msdn.microsoft.com/en-us/library/windows/desktop/ee886409(v=vs.85).aspx
	http://msdn.microsoft.com/en-us/library/system.security.principal.securityidentifier.aspx

.NOTES
	Author:	 Brian C. Wilhite
	Email:	 bwilhite1@carolina.rr.com
	Date: 	 "09/20/2012"
	Updates: Added FilterSID Parameter
	         Cleaned Up Code, defined fewer variables when creating PSObjects
	ToDo:    Clean up the UserSID Translation, to continue even if the SID is local
#>

[CmdletBinding()]
param(
	[Parameter(Position=0,ValueFromPipeline=$true)]
	[Alias("CN","Computer")]
	[String[]]$ComputerName="$env:COMPUTERNAME",
	[String]$FilterSID,
	[String]$WQLFilter="NOT SID = 'S-1-5-18' AND NOT SID = 'S-1-5-19' AND NOT SID = 'S-1-5-20'"
	)

Begin
	{
		#Adjusting ErrorActionPreference to stop on all errors
		$TempErrAct = $ErrorActionPreference
		$ErrorActionPreference = "Stop"
		#Exclude Local System, Local Service & Network Service
	}#End Begin Script Block

Process
	{
		Foreach ($Computer in $ComputerName)
			{
				$Computer = $Computer.ToUpper().Trim()
				Try
					{
						#Querying Windows version to determine how to proceed.
						$Win32OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer
						$Build = $Win32OS.BuildNumber
						
						#Win32_UserProfile exist on Windows Vista and above
						If ($Build -ge 6001)
							{
								If ($FilterSID)
									{
										$WQLFilter = $WQLFilter + " AND NOT SID = `'$FilterSID`'"
									}#End If ($FilterSID)
								$Win32User = Get-WmiObject -Class Win32_UserProfile -Filter $WQLFilter -ComputerName $Computer
								$LastUser = $Win32User | Sort-Object -Property LastUseTime -Descending | Select-Object -First 1
								$Loaded = $LastUser.Loaded
								$Script:Time = ([WMI]'').ConvertToDateTime($LastUser.LastUseTime)
								
								#Convert SID to Account for friendly display
								$Script:UserSID = New-Object System.Security.Principal.SecurityIdentifier($LastUser.SID)
								$User = $Script:UserSID.Translate([System.Security.Principal.NTAccount])
							}#End If ($Build -ge 6001)
							
						If ($Build -le 6000)
							{
								If ($Build -eq 2195)
									{
										$SysDrv = $Win32OS.SystemDirectory.ToCharArray()[0] + ":"
									}#End If ($Build -eq 2195)
								Else
									{
										$SysDrv = $Win32OS.SystemDrive
									}#End Else
								$SysDrv = $SysDrv.Replace(":","$")
								$Script:ProfLoc = "\\$Computer\$SysDrv\Documents and Settings"
								$Profiles = Get-ChildItem -Path $Script:ProfLoc
								$Script:NTUserDatLog = $Profiles | ForEach-Object -Process {$_.GetFiles("ntuser.dat.LOG")}
								
								#Function to grab last profile data, used for allowing -FilterSID to Function properly.
								Function GetLastProfData ($InstanceNumber)
									{
										$Script:LastProf = ($Script:NTUserDatLog | Sort-Object -Property LastWriteTime -Descending)[$InstanceNumber]							
										$Script:UserName = $Script:LastProf.DirectoryName.Replace("$Script:ProfLoc","").Trim("\").ToUpper()
										$Script:Time = $Script:LastProf.LastAccessTime
										
										#Getting the SID of the user from the file ACE to compare
										$Script:Sddl = $Script:LastProf.GetAccessControl().Sddl
										$Script:Sddl = $Script:Sddl.split("(") | Select-String -Pattern "[0-9]\)$" | Select-Object -First 1
										#Formatting SID, assuming the 6th entry will be the users SID.
										$Script:Sddl = $Script:Sddl.ToString().Split(";")[5].Trim(")")
										
										#Convert Account to SID to detect if profile is loaded via the remote registry
										$Script:TranSID = New-Object System.Security.Principal.NTAccount($Script:UserName)
										$Script:UserSID = $Script:TranSID.Translate([System.Security.Principal.SecurityIdentifier])
									}#End Function GetLastProfData
								GetLastProfData -InstanceNumber 0
								
								#If the FilterSID equals the UserSID, rerun GetLastProfData and select the next instance
								If ($Script:UserSID -eq $FilterSID)
									{
										GetLastProfData -InstanceNumber 1
									}#End If ($Script:UserSID -eq $FilterSID)
								
								#If the detected SID via Sddl matches the UserSID, then connect to the registry to detect currently loggedon.
								If ($Script:Sddl -eq $Script:UserSID)
									{
										$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]"Users",$Computer)
										$Loaded = $Reg.GetSubKeyNames() -contains $Script:UserSID.Value
										#Convert SID to Account for friendly display
										$Script:UserSID = New-Object System.Security.Principal.SecurityIdentifier($Script:UserSID)
										$User = $Script:UserSID.Translate([System.Security.Principal.NTAccount])
									}#End If ($Script:Sddl -eq $Script:UserSID)
								Else
									{
										$User = $Script:UserName
										$Loaded = "Unknown"
									}#End Else

							}#End If ($Build -le 6000)
						
						#Creating Custom PSObject For Output
						New-Object -TypeName PSObject -Property @{
							Computer=$Computer
							User=$User
							SID=$Script:UserSID
							Time=$Script:Time
							CurrentlyLoggedOn=$Loaded
							} | Select-Object Computer, User, SID, Time, CurrentlyLoggedOn
							
					}#End Try
					
				Catch
					{
						If ($_.Exception.Message -Like "*Some or all identity references could not be translated*")
							{
								Write-Warning "Unable to Translate $Script:UserSID, try filtering the SID `nby using the -FilterSID parameter."	
								Write-Warning "It may be that $Script:UserSID is local to $Computer, Unable to translate remote SID"
							}
						Else
							{
								Write-Warning $_
							}
					}#End Catch
					
			}#End Foreach ($Computer in $ComputerName)
			
	}#End Process
	
End
	{
		#Resetting ErrorActionPref
		$ErrorActionPreference = $TempErrAct
	}#End End

}# End Function Get-LastLogon
Function Get-LHSAntiVirusProduct 
{
<#
.SYNOPSIS
    Get the status of Antivirus Product on local and Remote Computers.

.DESCRIPTION
    It works with MS Security Center and detects the status for most AV products.
    
    Note that this script will only work on Windows XP SP2, Vista, 7, 8.x, 10 
    operating systems as Windows Servers does not have 
    the required WMI SecurityCenter\SecurityCenter(2) name spaces.

.PARAMETER ComputerName
    The computer name(s) to retrieve the info from. 

.EXAMPLE
    Get-LHSAntiVirusProduct
    
    ComputerName             : Localhost
    Name                     : Kaspersky Endpoint Security 10 für Windows
    ProductExecutable        : C:\Program Files (x86)\Kaspersky Lab\Kaspersky Endpoint 
                               Security 10 for Windows SP1\wmiav.exe
    DefinitionStatus         : UP_TO_DATE
    RealTimeProtectionStatus : ON
    ProductState             : 266240
 
.EXAMPLE
    Get-LHSAntiVirusProduct –ComputerName PC1,PC2,PC3

    ComputerName             : PC1
    Name                     : Kaspersky Endpoint Security 10 für Windows
    ProductExecutable        : C:\Program Files (x86)\Kaspersky Lab\Kaspersky Endpoint 
                               Security 10 for Windows SP1\wmiav.exe
    DefinitionStatus         : UP_TO_DATE
    RealTimeProtectionStatus : ON
    ProductState             : 266240
    (..)

.EXAMPLE
    (get-content PClist.txt) | Get-LHSAntiVirusProduct

 .INPUTS
    System.String, you can pipe ComputerNames to this Function

.OUTPUTS
    Custom PSObjects 

.NOTE
    WMI query to get anti-virus infor­ma­tion has been changed.
    Pre-Vista clients used the root/SecurityCenter name­space, 
    while Post-Vista clients use the root/SecurityCenter2 name­space.
    But not only the name­space has been changed, The properties too. 


    More info at http://neophob.com/2010/03/wmi-query-windows-securitycenter2/
    and from this MSDN Blog 
    http://blogs.msdn.com/b/alejacma/archive/2008/05/12/how-to-get-antivirus-information-with-wmi-vbscript.aspx


    AUTHOR: Pasquale Lantella 
    LASTEDIT: 23.06.2016
    KEYWORDS: Antivirus
    Version :1.1
    History :1.1 support for Win 10, changed the use of WMI productState   

.LINK
    WSC_SECURITY_PRODUCT_STATE enumeration
    https://msdn.microsoft.com/en-us/library/jj155490%28v=vs.85%29

.LINK
    Windows Security Center
    https://msdn.microsoft.com/en-us/library/gg537273%28v=vs.85%29

.LINK
    http://neophob.com/2010/03/wmi-query-windows-securitycenter2/

#Requires -Version 2.0
#>


[CmdletBinding()]

param (
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias('CN')]
    [String[]]$ComputerName=$env:computername
)

BEGIN {

    Set-StrictMode -Version Latest
    ${CmdletName} = $Pscmdlet.MyInvocation.MyCommand.Name

} # end BEGIN

PROCESS {
    
    ForEach ($Computer in $computerName) 
    {
        IF (Test-Connection -ComputerName $Computer -count 2 -quiet) 
        { 
            Try
            {
                [system.Version]$OSVersion = (Get-WmiObject win32_operatingsystem -computername $Computer).version

                IF ($OSVersion -ge [system.version]'6.0.0.0') 
                {
                    Write-Verbose "OS Windows Vista/Server 2008 or newer detected."
                    $AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct -ComputerName $Computer -ErrorAction Stop
                } 
                Else 
                {
                    Write-Verbose "Windows 2000, 2003, XP detected" 
                    $AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter -Class AntiVirusProduct  -ComputerName $Computer -ErrorAction Stop
                } # end IF ($OSVersion -ge 6.0) 
 
                <#
                it appears that if you convert the productstate to HEX then you can read the 1st 2nd or 3rd block 
                to get whether product is enabled/disabled and whether definitons are up-to-date or outdated
                #>

                $productState = $AntiVirusProduct.productState

                # convert to hex, add an additional '0' left if necesarry
                $hex = [Convert]::ToString($productState, 16).PadLeft(6,'0')

                # Substring(int startIndex, int length)  
                $WSC_SECURITY_PROVIDER = $hex.Substring(0,2)
                $WSC_SECURITY_PRODUCT_STATE = $hex.Substring(2,2)
                $WSC_SECURITY_SIGNATURE_STATUS = $hex.Substring(4,2)

                #n ot used yet
                $SECURITY_PROVIDER = switch ($WSC_SECURITY_PROVIDER)
                {
                    0  {"NONE"}
                    1  {"FIREWALL"}
                    2  {"AUTOUPDATE_SETTINGS"}
                    4  {"ANTIVIRUS"}
                    8  {"ANTISPYWARE"}
                    16 {"INTERNET_SETTINGS"}
                    32 {"USER_ACCOUNT_CONTROL"}
                    64 {"SERVICE"}
                    default {"UNKNOWN"}
                }


                $RealTimeProtectionStatus = switch ($WSC_SECURITY_PRODUCT_STATE)
                {
                    "00" {"OFF"} 
                    "01" {"EXPIRED"}
                    "10" {"ON"}
                    "11" {"SNOOZED"}
                    default {"UNKNOWN"}
                }

                $DefinitionStatus = switch ($WSC_SECURITY_SIGNATURE_STATUS)
                {
                    "00" {"UP_TO_DATE"}
                    "10" {"OUT_OF_DATE"}
                    default {"UNKNOWN"}
                }  

<#  
                # Switch to determine the status of antivirus definitions and real-time protection.
                # The values in this switch-statement are retrieved from the following website: http://community.kaseya.com/resources/m/knowexch/1020.aspx
                switch ($AntiVirusProduct.productState) {
                     #AVG Internet Security 2012 (from antivirusproduct WMI)
                     "262144" {$defstatus = "Up to date" ;$rtstatus = "Disabled"}
                     "266240" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}
 
                     "262160" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                     "266256" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                     "393216" {$defstatus = "Up to date" ;$rtstatus = "Disabled"}
                     "393232" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                     "393488" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                     "397312" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}
                     "397328" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                     #Windows Defender
                     "393472" {$defstatus = "Up to date" ;$rtstatus = "Disabled"} 
                     "397584" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                     "397568" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}

                     default {$defstatus = "Unknown" ;$rtstatus = "Unknown"}
                }
#>

              
                # Output PSCustom Object
                $AV = $Null
                $AV = New-Object -TypeName PSObject -ErrorAction Stop -Property @{
             
                    ComputerName = $AntiVirusProduct.__Server;
                    Name = $AntiVirusProduct.displayName;
                    ProductExecutable = $AntiVirusProduct.pathToSignedProductExe;
                    DefinitionStatus = $DefinitionStatus;
                    RealTimeProtectionStatus = $RealTimeProtectionStatus;
                    ProductState = $productState;
                
                } | Select-Object ComputerName,Name,ProductExecutable,DefinitionStatus,RealTimeProtectionStatus,ProductState  
                
                Write-Output $AV 
            }
            Catch 
            {
                Write-Error "\\$Computer : WMI Error"
                Write-Error $_
            }                              
        } 
        Else 
        {
            Write-Warning "\\$computer DO NOT reply to ping" 
        } # end IF (Test-Connection -ComputerName $Computer -count 2 -quiet)
	   
    } # end ForEach ($Computer in $computerName)

} # end PROCESS

END { Write-Verbose "Function Get-LHSAntiVirusProduct finished." } 
} # end Function Get-LHSAntiVirusProduct


Function Get-LocalAdministratorBuiltin
{
<#
	.SYNOPSIS
		Function to retrieve the local Administrator account
	
	.DESCRIPTION
		Function to retrieve the local Administrator account
	
	.PARAMETER ComputerName
		Specifies the computername
	
	.EXAMPLE
		PS C:\> Get-LocalAdministratorBuiltin
	
	.EXAMPLE
		PS C:\> Get-LocalAdministratorBuiltin -ComputerName SERVER01
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
	
	#Function to get the BUILTIN LocalAdministrator
	#http://blog.simonw.se/powershell-find-builtin-local-administrator-account/
#>
	
	[CmdletBinding()]
	param (
		[Parameter()]
		$ComputerName = $env:computername
	)
	Process
	{
		Foreach ($Computer in $ComputerName)
		{
			Try
			{
				Add-Type -AssemblyName System.DirectoryServices.AccountManagement
				$PrincipalContext = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine, $Computer)
				$UserPrincipal = New-Object -TypeName System.DirectoryServices.AccountManagement.UserPrincipal($PrincipalContext)
				$Searcher = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalSearcher
				$Searcher.QueryFilter = $UserPrincipal
				$Searcher.FindAll() | Where-Object { $_.Sid -Like "*-500" }
			}
			Catch
			{
				Write-Warning -Message "$($_.Exception.Message)"
			}
		}
	}
}
Function Get-LocalGroup
{
	
<#
	.SYNOPSIS
		This script can be list all of local group account.
	
	.DESCRIPTION
		This script can be list all of local group account.
		The Function is using WMI to connect to the remote machine
	
	.PARAMETER ComputerName
		Specifies the computers on which the command . The default is the local computer.
	
	.PARAMETER Credential
		A description of the Credential parameter.
	
	
	.EXAMPLE
		Get-LocalGroup
		
		This example shows how to list all the local groups on local computer.
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	
	PARAM
	(
		[Alias('cn')]
		[String[]]$ComputerName = $Env:COMPUTERNAME,
		
		[String]$AccountName,
		
		[System.Management.Automation.PsCredential]$Credential
	)
	
	$Splatting = @{
		Class = "Win32_Group"
		Namespace = "root\cimv2"
		Filter = "LocalAccount='$True'"
	}
	
	#Credentials
	If ($PSBoundParameters['Credential']) { $Splatting.Credential = $Credential }
	
	Foreach ($Computer in $ComputerName)
	{
		TRY
		{
			Write-Verbose -Message "[PROCESS] ComputerName: $Computer"
			Get-WmiObject @Splatting -ComputerName $Computer | Select-Object -Property Name, Caption, Status, SID, SIDType, Domain, Description
		}
		CATCH
		{
			Write-Warning -Message "[PROCESS] Issue connecting to $Computer"
		}
	}
}
#requires -version 4.0

Function Get-LocalGroupMember {

<#
.SYNOPSIS
Get local group membership using ADSI.

.DESCRIPTION
This command uses ADSI to connect to a server and enumerate the members of a local group. By default it will retrieve members of the local Administrators group.

The command uses legacy protocols to connect and enumerate group memberships. You may find it more efficient to wrap this Function in an Invoke-Command expression. See examples.

.PARAMETER Computername
The name of a computer to query. The parameter has aliases of 'CN' and 'Host'.

.PARAMETER Name
The name of a local group. 

.EXAMPLE
PS C:\> Get-LocalGroupMember -computer chi-core01

Computername : CHI-CORE01
Name         : Administrator
ADSPath      : WinNT://GLOBOMANTICS/chi-core01/Administrator
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : True

Computername : CHI-CORE01
Name         : Domain Admins
ADSPath      : WinNT://GLOBOMANTICS/Domain Admins
Class        : Group
Domain       : GLOBOMANTICS
IsLocal      : False

Computername : CHI-CORE01
Name         : Chicago IT
ADSPath      : WinNT://GLOBOMANTICS/Chicago IT
Class        : Group
Domain       : GLOBOMANTICS
IsLocal      : False

Computername : CHI-CORE01
Name         : OMAA
ADSPath      : WinNT://GLOBOMANTICS/OMAA
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : False

Computername : CHI-CORE01
Name         : LocalAdmin
ADSPath      : WinNT://GLOBOMANTICS/chi-core01/LocalAdmin
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : True

.EXAMPLE
PS C:\> "chi-hvr1","chi-hvr2","chi-core01","chi-fp02" | get-localgroupmember  | where {$_.IsLocal} | Select Computername,Name,ADSPath

Computername Name          ADSPath                                      
------------ ----          -------                                      
CHI-HVR1     Administrator WinNT://GLOBOMANTICS/chi-hvr1/Administrator  
CHI-HVR2     Administrator WinNT://GLOBOMANTICS/chi-hvr2/Administrator  
CHI-HVR2     Jeff          WinNT://GLOBOMANTICS/chi-hvr2/Jeff           
CHI-CORE01   Administrator WinNT://GLOBOMANTICS/chi-core01/Administrator
CHI-CORE01   LocalAdmin    WinNT://GLOBOMANTICS/chi-core01/LocalAdmin   
CHI-FP02     Administrator WinNT://GLOBOMANTICS/chi-fp02/Administrator

.EXAMPLE
PS C:\> $s = new-pssession chi-hvr1,chi-fp02,chi-hvr2,chi-core01
Create several PSSessions to remote computers.

PS C:\> $sb = ${Function:Get-localGroupMember} 

Get the Function's scriptblock

PS C:\> Invoke-Command -scriptblock { new-item -path Function:Get-LocalGroupMember -value $using:sb} -session $s 

Create a remote version of the Function.

PS C:\> Invoke-Command -scriptblock { get-localgroupmember | where {$_.IsLocal} } -session $s | Select Computername,Name,ADSPath

Repeat an example from above but this time execute it in a remote session.

.EXAMPLE
PS C:\> get-localgroupmember -Name "Hyper-V administrators" -Computername chi-hvr1,chi-hvr2


Computername : CHI-HVR1
Name         : jeff
ADSPath      : WinNT://GLOBOMANTICS/jeff
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : False

Computername : CHI-HVR2
Name         : jeff
ADSPath      : WinNT://GLOBOMANTICS/jeff
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : False

Check group membership for the Hyper-V Administrators group.

.EXAMPLE
PS C:\> get-localgroupmember -Computername chi-core01 | where class -eq 'group' | select Domain,Name

Domain       Name         
------       ----         
GLOBOMANTICS Domain Admins
GLOBOMANTICS Chicago IT   

Get members of the Administrators group on CHI-CORE01 that are groups and select a few properties.


.NOTES
NAME        :  Get-LocalGroupMember
VERSION     :  1.6   
LAST UPDATED:  2/18/2016
AUTHOR      :  Jeff Hicks (@JeffHicks)

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************

.INPUTS
[string] for computer names

.OUTPUTS
[object]

#>


[cmdletbinding()]

Param(
[Parameter(Position = 0)]
[ValidateNotNullorEmpty()]
[string]$Name = "Administrators",

[Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
[ValidateNotNullorEmpty()]
[Alias("CN","host")]
[string[]]$Computername = $env:computername
)


Begin {
    Write-Verbose "[Starting] $($MyInvocation.Mycommand)"  
    Write-Verbose "[Begin]    Querying members of the $Name group"
} #begin

Process {
 
 foreach ($computer in $computername) {

    #define a flag to indicate if there was an error
    $script:NotFound = $False
    
    #define a trap to handle errors because we're not using cmdlets that
    #could support Try/Catch. Traps must be in same scope.
    Trap [System.Runtime.InteropServices.COMException] {
        $errMsg = "Failed to enumerate $name on $computer. $($_.exception.message)"
        Write-Warning $errMsg

        #set a flag
        $script:NotFound = $True
    
        Continue    
    }

    #define a Trap for all other errors
    Trap {
      Write-Warning "Oops. There was some other type of error: $($_.exception.message)"
      Continue
    }

    Write-Verbose "[Process]  Connecting to $computer"
    #the WinNT moniker is case-sensitive
    [ADSI]$group = "WinNT://$computer/$Name,group"
        
    Write-Verbose "[Process]  Getting group member details" 
    $members = $group.invoke("Members") 

    Write-Verbose "[Process]  Counting group members"
    
    if (-Not $script:NotFound) {
        $found = ($members | measure).count
        Write-Verbose "[Process]  Found $found members"

        if ($found -gt 0 ) {
        $members | foreach {
        
            #define an ordered hashtable which will hold properties
            #for a custom object
            $Hash = [ordered]@{Computername = $computer.toUpper()}

            #Get the name property
            $hash.Add("Name",$_[0].GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null))
        
            #get ADS Path of member
            $ADSPath = $_[0].GetType().InvokeMember("ADSPath", 'GetProperty', $null, $_, $null)
            $hash.Add("ADSPath",$ADSPath)
    
            #get the member class, ie user or group
            $hash.Add("Class",$_[0].GetType().InvokeMember("Class", 'GetProperty', $null, $_, $null))  
    
            <#
            Domain members will have an ADSPath like WinNT://MYDomain/Domain Users.  
            Local accounts will be like WinNT://MYDomain/Computername/Administrator
            #>

            $hash.Add("Domain",$ADSPath.Split("/")[2])

            #if computer name is found between two /, then assume
            #the ADSPath reflects a local object
            if ($ADSPath -match "/$computer/") {
                $local = $True
                }
            else {
                $local = $False
                }
            $hash.Add("IsLocal",$local)

            #turn the hashtable into an object
            New-Object -TypeName PSObject -Property $hash
         } #foreach member
        } 
        else {
            Write-Warning "No members found in $Name on $Computer."
        }
    } #if no errors
} #foreach computer

} #process

End {
    Write-Verbose "[Ending]  $($MyInvocation.Mycommand)"
} #end

} #end Function

<#
Copyright (c) 2016 JDH Information Technology Solutions, Inc.


Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:


The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.


THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
#>
<#
===================================================================================  
DESCRIPTION:    Function enumerates members of all local groups (or a given group). 
If -Server parameter is not specified, it will query localhost by default. 
If -Group parameter is not specified, all local groups will be queried. 
            
AUTHOR:    	Piotr Lewandowski 
VERSION:    1.0  
DATE:       29/04/2013  
SYNTAX:     Get-LocalGroupMembers [[-Server] <string[]>] [[-Group] <string[]>] 
             
EXAMPLES:   

Get-LocalGroupMembers -server "scsm-server" | ft -AutoSize

Server      Local Group          Name                 Type  Domain  SID
------      -----------          ----                 ----  ------  ---
scsm-server Administrators       Administrator        User          S-1-5-21-1473970658-40817565-21663372-500
scsm-server Administrators       Domain Admins        Group contoso S-1-5-21-4081441239-4240563405-729182456-512
scsm-server Guests               Guest                User          S-1-5-21-1473970658-40817565-21663372-501
scsm-server Remote Desktop Users pladmin              User  contoso S-1-5-21-4081441239-4240563405-729182456-1272
scsm-server Users                INTERACTIVE          Group         S-1-5-4
scsm-server Users                Authenticated Users  Group         S-1-5-11



"scsm-dc01","scsm-server" | Get-LocalGroupMembers -group administrators | ft -autosize

Server      Local Group    Name                 Type  Domain  SID
------      -----------    ----                 ----  ------  ---
scsm-dc01   administrators Administrator        User  contoso S-1-5-21-4081441239-4240563405-729182456-500
scsm-dc01   administrators Enterprise Admins    Group contoso S-1-5-21-4081441239-4240563405-729182456-519
scsm-dc01   administrators Domain Admins        Group contoso S-1-5-21-4081441239-4240563405-729182456-512
scsm-server administrators Administrator        User          S-1-5-21-1473970658-40817565-21663372-500
scsm-server administrators !svcServiceManager   User  contoso S-1-5-21-4081441239-4240563405-729182456-1274
scsm-server administrators !svcServiceManagerWF User  contoso S-1-5-21-4081441239-4240563405-729182456-1275
scsm-server administrators !svcscoservice       User  contoso S-1-5-21-4081441239-4240563405-729182456-1310
scsm-server administrators Domain Admins        Group contoso S-1-5-21-4081441239-4240563405-729182456-512
 
===================================================================================  

#>
Function Get-LocalGroupMembers
{
param(
[Parameter(ValuefromPipeline=$true)][array]$server = $env:computername,
$GroupName = $null
)
PROCESS {
    $finalresult = @()
    $computer = [ADSI]"WinNT://$server"

    if (!($groupName))
    {
    $Groups = $computer.psbase.Children | Where {$_.psbase.schemaClassName -eq "group"} | select -expand name
    }
    else
    {
    $groups = $groupName
    }
    $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().GetDirectoryEntry() | select name,objectsid
    $domain = $currentdomain.name
    $SID=$CurrentDomain.objectsid
    $DomainSID = (New-Object System.Security.Principal.SecurityIdentifier($sid[0], 0)).value


    foreach ($group in $groups)
    {
    $gmembers = $null
    $LocalGroup = [ADSI]("WinNT://$server/$group,group")


    $GMembers = $LocalGroup.psbase.invoke("Members")
    $GMemberProps = @{Server="$server";"Local Group"=$group;Name="";Type="";ADSPath="";Domain="";SID=""}
    $MemberResult = @()


        if ($gmembers)
        {
        foreach ($gmember in $gmembers)
            {
            $membertable = new-object psobject -Property $GMemberProps
            $name = $gmember.GetType().InvokeMember("Name",'GetProperty', $null, $gmember, $null)
            $sid = $gmember.GetType().InvokeMember("objectsid",'GetProperty', $null, $gmember, $null)
            $UserSid = New-Object System.Security.Principal.SecurityIdentifier($sid, 0)
            $class = $gmember.GetType().InvokeMember("Class",'GetProperty', $null, $gmember, $null)
            $ads = $gmember.GetType().InvokeMember("adspath",'GetProperty', $null, $gmember, $null)
            $MemberTable.name= "$name"
            $MemberTable.type= "$class"
            $MemberTable.adspath="$ads"
            $membertable.sid=$usersid.value
            

            if ($userSID -like "$domainsid*")
                {
                $MemberTable.domain = "$domain"
                }
            
            $MemberResult += $MemberTable
            }
            
         }
         $finalresult += $MemberResult 
    }
    $finalresult | select server,"local group",name,type,domain,sid
    }
}
# ============================================================================================== 
# NAME: Listing Administrators and PowerUsers on remote machines  
#  
# AUTHOR: Mohamed Garrana ,  
# DATE  : 09/04/2010 
#  
# COMMENT:  
# This script runs against an input file of computer names , connects to each computer and gets a list of the users in the  local Administrators  
#and powerusers Groups . the output can be a csv file which can be readable on excel with all the computers from the input file 
# ============================================================================================== 
Function Get-LocalServerAdmins { 
        param( 
    [Parameter(Mandatory=$true,valuefrompipeline=$true)] 
    [string]$strComputer) 
    begin {} 
    Process { 
        $adminlist ="" 
        #$powerlist ="" 
        $computer = [ADSI]("WinNT://" + $strComputer + ",computer") 
        $AdminGroup = $computer.psbase.children.find("Administrators") 
        #$powerGroup = $computer.psbase.children.find("Power Users") 
        $Adminmembers= $AdminGroup.psbase.invoke("Members") | %{$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 
        #$Powermembers= $PowerGroup.psbase.invoke("Members") | %{$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 
        foreach ($admin in $Adminmembers) { $adminlist = $adminlist + $admin + "," } 
        #foreach ($poweruser in $Powermembers) { $powerlist = $powerlist + $poweruser + "," } 
        $Computer = New-Object psobject 
        $computer | Add-Member noteproperty ComputerName $strComputer 
        $computer | Add-Member noteproperty Administrators $adminlist 
        #$computer | Add-Member noteproperty PowerUsers $powerlist 
        Write-Output $computer 
 
 
        } 
end {} 
} 
 
#Get-Content C:\LazyWinAdmin\Servers\RESULTS\Alive\All.txt | Get-LocalServerAdmins | Export-Csv 'C:\LazyWinAdmin\Local Admin Accounts\LocalServerAdmins.csv' -NoTypeInformation 
Function Get-LocalUser
{
	
<#
	.SYNOPSIS
		This script can be list all of local user account.
	
	.DESCRIPTION
		This script can be list all of local user account.
		The Function is using WMI to connect to the remote machine
	
	.PARAMETER ComputerName
		Specifies the computers on which the command . The default is the local computer.
	
	.PARAMETER Credential
		A description of the Credential parameter.
	
	
	.EXAMPLE
		Get-LocalUser
		
		This example shows how to list all of local users on local computer.
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	
	PARAM
	(
		[Alias('cn')]
		[String[]]$ComputerName = $Env:COMPUTERNAME,
		
		[String]$AccountName,
		
		[System.Management.Automation.PsCredential]$Credential
	)
	
	$Splatting = @{
		Class = "Win32_UserAccount"
		Namespace = "root\cimv2"
		Filter = "LocalAccount='$True'"
	}
	
	#Credentials
	If ($PSBoundParameters['Credential']) { $Splatting.Credential = $Credential }
	
	Foreach ($Computer in $ComputerName)
	{
		TRY
		{
			Write-Verbose -Message "[PROCESS] ComputerName: $Computer"
			Get-WmiObject @Splatting -ComputerName $Computer | Select-Object -Property Name, FullName, Caption, Disabled, Status, Lockout, PasswordChangeable, PasswordExpires, PasswordRequired, SID, SIDType, AccountType, Domain, Description
		}
		CATCH
		{
			Write-Warning -Message "[PROCESS] Issue connecting to $Computer"
		}
	}
}
Function Get-LogFast
{
<#
    .DESCRIPTION
		Function to read a log file very fast
	.SYNOPSIS
		Function to read a log file very fast
	.EXAMPLE
		Get-LogFast -Path C:\megalogfile.log
    .EXAMPLE
        Get-LogFast -Path C:\367.msp.0.log -Match "09:36:43:417" -Verbose

        VERBOSE: [PROCESS] Match found
        MSI (s) (A8:14) [09:36:43:417]: Note: 1: 2205 2:  3: Font 
        VERBOSE: [PROCESS] Match found
        MSI (s) (A8:14) [09:36:43:417]: Note: 1: 2205 2:  3: Class 
        VERBOSE: [PROCESS] Match found
        MSI (s) (A8:14) [09:36:43:417]: Note: 1: 2205 2:  3: TypeLib 

	.NOTES
		Francois-Xavier cat
		@lazywinadm
		www.lazywinadmin.com
		github.com/lazywinadmin
		
#>
	[CmdletBinding()]
	PARAM (
		$Path = "c:\Biglog.log",
		
		$Match
	)
	BEGIN
	{
		# Create a StreamReader object
		#  Fortunately this .NET Framework called System.IO.StreamReader allows you to read text files a line at a time which is important when you’ re dealing with huge log files :-)
		$StreamReader = New-object -TypeName System.IO.StreamReader -ArgumentList (Resolve-Path -Path $Path -ErrorAction Stop).Path
	}
	PROCESS
	{
		# .Peek() Method: An integer representing the next character to be read, or -1 if no more characters are available or the stream does not support seeking.
		while ($StreamReader.Peek() -gt -1)
		{
			# Read the next line
			#  .ReadLine() method: Reads a line of characters from the current stream and returns the data as a string.
			$Line = $StreamReader.ReadLine()
			
			#  Ignore empty line and line starting with a #
			if ($Line.length -eq 0 -or $Line -match "^#")
			{
				continue
			}
			
			IF ($PSBoundParameters['Match'])
			{
				If ($Line -match $Match)
				{
					Write-Verbose -Message "[PROCESS] Match found"
					
					# Split the line on $Delimiter
					#$result = ($Line -split $Delimiter)
					
					Write-Output $Line
				}
			}
			ELSE { Write-Output $Line }
		}
	} #PROCESS
}
Function get-loggedonuser ($computername){

#mjolinor 3/17/10

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
<#
.Synopsis
   A quick Function to determine if a computer is VM or physical box.
.DESCRIPTION
   This Function is designed to quickly determine if a local or remote
   computer is a physical machine or a virtual machine.
.NOTES
   Created by: Jason Wasser
   Modified: 4/20/2017 03:28:53 PM  

   Changelog: 
    * Code cleanup thanks to suggestions from @juneb_get_help
    * added credential support
    * Added Xen AWS Xen for HVM domU

   To Do:
    * Find the Model information for other hypervisor VM's (i.e KVM).
.EXAMPLE
   Get-MachineType
   Query if the local machine is a physical or virtual machine.
.EXAMPLE
   Get-MachineType -ComputerName SERVER01 
   Query if SERVER01 is a physical or virtual machine.
.EXAMPLE
   Get-MachineType -ComputerName (Get-Content c:\temp\computerlist.txt)
   Query if a list of computers are physical or virtual machines.
.LINK
   https://gallery.technet.microsoft.com/scriptcenter/Get-MachineType-VM-or-ff43f3a9
#>
Function Get-MachineType
{
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param
    (
        # ComputerName
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string[]]$ComputerName=$env:COMPUTERNAME,
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin
    {
    }
    Process
    {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Checking $Computer"
            try {
                # Check to see if $Computer resolves DNS lookup successfuly.
                $null = [System.Net.DNS]::GetHostEntry($Computer)
                
                $ComputerSystemInfo = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop -Credential $Credential
                
                switch ($ComputerSystemInfo.Model) {
                    
                    # Check for Hyper-V Machine Type
                    "Virtual Machine" {
                        $MachineType="VM"
                        }

                    # Check for VMware Machine Type
                    "VMware Virtual Platform" {
                        $MachineType="VM"
                        }

                    # Check for Oracle VM Machine Type
                    "VirtualBox" {
                        $MachineType="VM"
                        }

                    # Check for Xen
                    "HVM domU" {
                        $MachineType="VM"
                        }

                    # Check for KVM
                    # I need the values for the Model for which to check.

                    # Otherwise it is a physical Box
                    default {
                        $MachineType="Physical"
                        }
                    }
                
                # Building MachineTypeInfo Object
                $MachineTypeInfo = New-Object -TypeName PSObject -Property ([ordered]@{
                    ComputerName=$ComputerSystemInfo.PSComputername
                    Type=$MachineType
                    Manufacturer=$ComputerSystemInfo.Manufacturer
                    Model=$ComputerSystemInfo.Model
                    })
                $MachineTypeInfo
                }
            catch [Exception] {
                Write-Output "$Computer`: $($_.Exception.Message)"
                }
            }
    }
    End
    {

    }
}
#Function Get-MappedDrives($ComputerName){
  #Ping remote machine, continue if available
  $ComputerName = Get-Content -Path C:\LazyWinAdmin\Logs\Contoso.corp\Computers\Windows7.txt
  if(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet){
    #Get remote explorer session to identify current user
    $explorer = Get-WmiObject -ComputerName $ComputerName -Class win32_process | ?{$_.name -eq "explorer.exe"}
    
    #If a session was returned check HKEY_USERS for Network drives under their SID
    if($explorer){
      $Hive = [long]$HIVE_HKU = 2147483651
      $sid = ($explorer.GetOwnerSid()).sid
      $owner  = $explorer.GetOwner()
      $RegProv = get-WmiObject -List -Namespace "root\default" -ComputerName $ComputerName | Where-Object {$_.Name -eq "StdRegProv"}
      $DriveList = $RegProv.EnumKey($Hive, "$($sid)\Network")
      
      #If the SID network has mapped drives iterate and report on said drives
      if($DriveList.sNames.count -gt 0){
        "$($owner.Domain)\$($owner.user) on $($ComputerName)"
        foreach($drive in $DriveList.sNames){
          "$($drive)`t$(($RegProv.GetStringValue($Hive, "$($sid)\Network\$($drive)", "RemotePath")).sValue)"
        }
      }else{"No mapped drives on $($ComputerName)"}
    }else{"explorer.exe not running on $($ComputerName)"}
  }else{"Can't connect to $($ComputerName)"}

Out-File -FilePath "C:\LazyWinAdmin\Logs\Contoso.corp\Computers\Mapped\" + $owner + ".txt"

Function Get-RebootBoolean
{
    Param
    (
        $ComputerName
    )
    Process
    {
        
        $os = Get-WmiObject win32_operatingsystem -ComputerName $ComputerName
        $uptime = (Get-Date) - ($os.ConvertToDateTime($os.lastbootuptime))
        $minutesUp=$uptime.TotalMinutes
   
    }
    End
    {
        if($minutesUp -le 120){
            return $true
        }else{
            return $false
        }
         
    }
}
Function Get-ProcessorBoolean
{
    Param
    (
        $ComputerName
    )

    Begin
    {
    }
    Process
    {
        $value=(Get-Counter -ComputerName $ComputerName -Counter “\Processor(_Total)\% Processor Time” -SampleInterval 10).CounterSamples.CookedValue
    }
    End
    {
        if($value -ge 90){
        return $true
        }else{
        return $false
        }
    }
}
Function Get-MemoryBoolean
{
    Param
    (
        $ComputerName
    )

    Process
    {
        $value=gwmi -Class win32_operatingsystem -computername $ComputerName | Select-Object @{Name = "MemoryUsage"; Expression = {“{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}
    }
    End
    {
        if($value.MemoryUsage -ge 90){
            return $true
        }else{
            return $false
        }
        
    }
}

Function Get-DiskSpaceBoolean
{
    Param
    (
        $freeBoolean=$false,
        $ComputerName
    )

    Process
    {
        $diskInfo=Get-WmiObject -ComputerName $ComputerName -class win32_logicaldisk
        foreach($disk in $diskInfo){
            if($disk.DeviceID -ne 'A:'){
                if(($disk.FreeSpace/$disk.Size)*100 -le 10){
                    $freeBoolean=$true
                }
            }

        }
    }
    End
    {
        $freeBoolean
    }
}


Function Get-NotRunningServices
{
    
    Param
    (
        $ComputerName
    )

    
    Process
    {
        $notRunning=Get-wmiobject -ComputerName $ComputerName win32_service -Filter "startmode = 'auto' AND state != 'running' AND Exitcode !=0"
        $count=$notRunning.Count
    }
    End
    {
        if($count -ge 0){
            return $true
        }
        else{
            return $false
        }
    }
}
Function Get-MOTD {

<#
.NAME
    Get-MOTD
.SYNOPSIS
    Displays system information to a host.
.DESCRIPTION
    The Get-MOTD cmdlet is a system information tool written in PowerShell. 
.EXAMPLE
#>


  [CmdletBinding()]
	
  Param(
    [Parameter(Position=0,Mandatory=$false)]
	[ValidateNotNullOrEmpty()]
    [string[]]$ComputerName
    ,
    [Parameter(Position=1,Mandatory=$false)]
    [PSCredential]
    [System.Management.Automation.CredentialAttribute()]$Credential
  )

  Begin {
	
        If (-Not $ComputerName) {
            $RemoteSession = $null
        }
        #Define ScriptBlock for data collection
        $ScriptBlock = {
            $Operating_System = Get-CimInstance -ClassName Win32_OperatingSystem
            $Logical_Disk = Get-CimInstance -ClassName Win32_LogicalDisk |
            Where-Object -Property DeviceID -eq $Operating_System.SystemDrive
			Try {
				$PCLi = Get-PowerCLIVersion
				$PCLiVer = ' | PowerCLi ' + [string]$PCLi.Major + '.' + [string]$PCLi.Minor + '.' + [string]$PCLi.Revision + '.' + [string]$PCLi.Build
			} Catch {$PCLiVer = ''}
			If ($DomainName = ([System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()).DomainName) {$DomainName = '.' + $DomainName}
			
            [pscustomobject]@{
                Operating_System = $Operating_System
                Processor = Get-CimInstance -ClassName Win32_Processor
                Process_Count = (Get-Process).Count
                Shell_Info = ("{0}.{1}" -f $PSVersionTable.PSVersion.Major,$PSVersionTable.PSVersion.Minor) + $PCLiVer
                Logical_Disk = $Logical_Disk
            }
        }
  } #End Begin

  Process {
	
        If ($ComputerName) {
            If ("$ComputerName" -ne "$env:ComputerName") {
                # Build Hash to be used for passing parameters to 
                # New-PSSession commandlet
                $PSSessionParams = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }

                # Add optional parameters to hash
                If ($Credential) {
                    $PSSessionParams.Add('Credential', $Credential)
                }

                # Create remote powershell session   
                Try {
                    $RemoteSession = New-PSSession @PSSessionParams
                }
                Catch {
                    Throw $_.Exception.Message
                }
            } Else { 
                $RemoteSession = $null
            }
        }
        
        # Build Hash to be used for passing parameters to 
        # Invoke-Command commandlet
        $CommandParams = @{
            ScriptBlock = $ScriptBlock
            ErrorAction = 'Stop'
        }
        
        # Add optional parameters to hash
        If ($RemoteSession) {
            $CommandParams.Add('Session', $RemoteSession)
        }
               
        # Run ScriptBlock    
        Try {
            $ReturnedValues = Invoke-Command @CommandParams
        }
        Catch {
            If ($RemoteSession) {
            	Remove-PSSession $RemoteSession
            }
            Throw $_.Exception.Message
        }

        # Assign variables
        Import-Module MS-Module
        $Date = Get-Date
        $OS_Name = $ReturnedValues.Operating_System.Caption + ' [Installed: ' + ([datetime]$ReturnedValues.Operating_System.InstallDate).ToString('dd-MMM-yyyy') + ']'
        $Computer_Name = $ReturnedValues.Operating_System.CSName
		If ($DomainName) {$Computer_Name = $Computer_Name + $DomainName.ToUpper()}
        $Kernel_Info = $ReturnedValues.Operating_System.Version + ' [' + $ReturnedValues.Operating_System.OSArchitecture + ']'
        $Process_Count = $ReturnedValues.Process_Count
        $Uptime = "$(($Uptime = $Date - $($ReturnedValues.Operating_System.LastBootUpTime)).Days) days, $($Uptime.Hours) hours, $($Uptime.Minutes) minutes"
        $Shell_Info = $ReturnedValues.Shell_Info
        $CPU_Info = $ReturnedValues.Processor.Name -replace '\(C\)', '' -replace '\(R\)', '' -replace '\(TM\)', '' -replace 'CPU', '' -replace '\s+', ' '
        $Current_Load = $ReturnedValues.Processor.LoadPercentage    
        $Memory_Size = "{0} MB/{1} MB " -f (([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))-
        ([math]::round($ReturnedValues.Operating_System.FreePhysicalMemory/1KB))),([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))
		$Disk_Size = "{0} GB/{1} GB" -f (([math]::round($ReturnedValues.Logical_Disk.Size/1GB)-
        [math]::round($ReturnedValues.Logical_Disk.FreeSpace/1GB))),([math]::round($ReturnedValues.Logical_Disk.Size/1GB))

        # Write to the Console
        Write-Host -Object ("")
        Write-Host -Object ("")
        Write-Host -Object ("         ,.=:^!^!t3Z3z.,                  ") -ForegroundColor Red
        Write-Host -Object ("        :tt:::tt333EE3                    ") -ForegroundColor Red
        Write-Host -Object ("        Et:::ztt33EEE ") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @Ee.,      ..,     $($Date.ToString('dd-MMM-yyyy HH:mm:ss'))") -ForegroundColor Green
        Write-Host -Object ("       ;tt:::tt333EE7") -NoNewline -ForegroundColor Red
        Write-Host -Object (" ;EEEEEEttttt33#     ") -ForegroundColor Green
        Write-Host -Object ("      :Et:::zt333EEQ.") -NoNewline -ForegroundColor Red
        Write-Host -Object (" SEEEEEttttt33QL     ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("User: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$env:USERDOMAIN\$env:UserName") -ForegroundColor Cyan
        Write-Host -Object ("      it::::tt333EEF") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @EEEEEEttttt33F      ") -NoNewline -ForeGroundColor Green
        Write-Host -Object ("Hostname: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Computer_Name") -ForegroundColor Cyan
        Write-Host -Object ("     ;3=*^``````'*4EEV") -NoNewline -ForegroundColor Red
        Write-Host -Object (" :EEEEEEttttt33@.      ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("OS: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$OS_Name") -ForegroundColor Cyan
        Write-Host -Object ("     ,.=::::it=., ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("``") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @EEEEEEtttz33QF       ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("Kernel: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("NT ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("$Kernel_Info") -ForegroundColor Cyan
        Write-Host -Object ("    ;::::::::zt33) ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("  '4EEEtttji3P*        ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("Uptime: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Uptime") -ForegroundColor Cyan
        Write-Host -Object ("   :t::::::::tt33.") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (":Z3z.. ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object (" ````") -NoNewline -ForegroundColor Green
        Write-Host -Object (" ,..g.        ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Shell: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("PowerShell $Shell_Info") -ForegroundColor Cyan
        Write-Host -Object ("   i::::::::zt33F") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" AEEEtttt::::ztF         ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("CPU: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$CPU_Info") -ForegroundColor Cyan
        Write-Host -Object ("  ;:::::::::t33V") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" ;EEEttttt::::t3          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Processes: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Process_Count") -ForegroundColor Cyan
        Write-Host -Object ("  E::::::::zt33L") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" @EEEtttt::::z3F          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Current Load: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Current_Load") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("%") -ForegroundColor Cyan
        Write-Host -Object (" {3=*^``````'*4E3)") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" ;EEEtttt:::::tZ``          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Memory: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Memory_Size`t") -ForegroundColor Cyan -NoNewline
		New-PercentageBar -DrawBar -Value (([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))-([math]::round($ReturnedValues.Operating_System.FreePhysicalMemory/1KB))) -MaxValue ([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB)); "`r"
        Write-Host -Object ("             ``") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" :EEEEtttt::::z7            ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("System Volume: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Disk_Size`t") -ForegroundColor Cyan -NoNewline
		New-PercentageBar -DrawBar -Value (([math]::round($ReturnedValues.Logical_Disk.Size/1GB)-[math]::round($ReturnedValues.Logical_Disk.FreeSpace/1GB))) -MaxValue ([math]::round($ReturnedValues.Logical_Disk.Size/1GB)); "`r"
        Write-Host -Object ("                 'VEzjt:;;z>*``           ") -ForegroundColor Yellow
        Write-Host -Object ("                      ````                  ") -ForegroundColor Yellow
        Write-Host -Object ("")
  } #End Process

  End {
        If ($RemoteSession) {
            Remove-PSSession $RemoteSession
        }
  }
} #End Function Get-MOTD
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
Function Get-Accelerators
{
	[psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::get
}
Function Get-NetFramework
{
	<#
	.SYNOPSIS
		This Function will retrieve the list of Framework Installed on the computer.
	.EXAMPLE
		Get-NetFramework
	
		PSChildName                                   Version                                      
		-----------                                   -------                                      
		v2.0.50727                                    2.0.50727.4927                               
		v3.0                                          3.0.30729.4926                               
		Windows Communication Foundation              3.0.4506.4926                                
		Windows Presentation Foundation               3.0.6920.4902                                
		v3.5                                          3.5.30729.4926                               
		Client                                        4.5.51641                                    
		Full                                          4.5.51641                                    
		Client                                        4.0.0.0        
	
	.NOTES
		TODO:
			Credential support
			ComputerName
				$hklm = 2147483650
				$key = "SOFTWARE\Microsoft\NET Framework Setup"
				$value = "NDP"
				Get-wmiobject -list "StdRegProv" -namespace root\default -computername . |
				Invoke-WmiMethod -name GetDWORDValue -ArgumentList $hklm,$key,$value | select uvalue

            #http://stackoverflow.com/questions/27375012/check-remote-wmi-and-remote-registry
	#>
	[CmdletBinding()]
	PARAM (
		[String[]]$ComputerName,
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)
	
	$Splatting = @{
		ComputerName = $ComputerName
	}
	
	if ($PSBoundParameters['Credential']) { $Splatting.credential = $Credential }
	
	Invoke-Command @Splatting -ScriptBlock {
		Write-Verbose -Message "$pscomputername"
		
		# Get the Net Framework Installed
		$netFramework = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -recurse |
		Get-ItemProperty -name Version -EA 0 |
		Where-Object { $_.PSChildName -match '^(?!S)\p{L}' } |
		Select-Object -Property PSChildName, Version
		
		# Prepare output
		$Properties = @{
			ComputerName = "$($env:Computername)$($env:USERDNSDOMAIN)"
			PowerShellVersion = $psversiontable.PSVersion.Major
			NetFramework = $netFramework
		}
		New-Object -TypeName PSObject -Property $Properties
	}
}
Function Get-NetFrameworkTypeAccelerator
{
<#
.SYNOPSIS 
	Function to retrieve the list of Type Accelerator available
.EXAMPLE
	Get-NetFrameworkTypeAccelerator
	
	Return the list of Type Accelerator available on your system
.EXAMPLE
	Get-Accelerator
	
	Return the list of Type Accelerator available on your system
	This is a Function alias created by [Alias()]
.NOTES
	Francois-Xavier Cat
	lazywinadmin.com
	@lazywinadm
	github.com/lazywinadmin
#>
	[Alias('Get-Acceletrator')]
	PARAM ()
	[System.Management.Automation.PSObject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::get
}
Function Get-NetStat
{
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
Function Get-NetworkLevelAuthentication
{
<#
.SYNOPSIS
	This Function will get the NLA setting on a local machine or remote machine

.DESCRIPTION
	This Function will get the NLA setting on a local machine or remote machine

.PARAMETER  ComputerName
	Specify one or more computer to query

.PARAMETER  Credential
	Specify the alternative credential to use. By default it will use the current one.

.EXAMPLE
	Get-NetworkLevelAuthentication
	
	This will get the NLA setting on the localhost

	ComputerName     : XAVIERDESKTOP
	NLAEnabled       : True
	TerminalName     : RDP-Tcp
	TerminalProtocol : Microsoft RDP 8.0
	Transport        : tcp	

.EXAMPLE
	Get-NetworkLevelAuthentication -ComputerName DC01
	
	This will get the NLA setting on the server DC01

	ComputerName     : DC01
	NLAEnabled       : True
	TerminalName     : RDP-Tcp
	TerminalProtocol : Microsoft RDP 8.0
	Transport        : tcp

.EXAMPLE
	Get-NetworkLevelAuthentication -ComputerName DC01, SERVER01 -verbose
	
	This will get the NLA setting on the servers DC01 and the SERVER01

.EXAMPLE
	Get-Content .\Computers.txt | Get-NetworkLevelAuthentication -verbose
	
	This will get the NLA setting for all the computers listed in the file Computers.txt
	
.EXAMPLE
	Get-NetworkLevelAuthentication -ComputerName (Get-Content -Path .\Computers.txt)
	
	This will get the NLA setting for all the computers listed in the file Computers.txt
	
.NOTES
	DATE	: 2014/04/01
	AUTHOR	: Francois-Xavier Cat
	WWW		: http://lazywinadmin.com
	Twitter	: @lazywinadm
	
	Article : http://lazywinadmin.com/2014/04/powershell-getset-network-level.html
	GitHub	: https://github.com/lazywinadmin/PowerShell
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the Function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				[pscustomobject][ordered]@{
					'ComputerName' = $NLAinfo.PSComputerName
					'NLAEnabled' = $NLAinfo.UserAuthenticationRequired -as [bool]
					'TerminalName' = $NLAinfo.TerminalName
					'TerminalProtocol' = $NLAinfo.TerminalProtocol
					'Transport' = $NLAinfo.transport
				}
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{
		
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}
##########################################
#Created by Nigel Tatschner on 22/07/2013#
##########################################
<#
	.SYNOPSIS
		 To gather Information about a Single or Multiple Systems.

	.DESCRIPTION
		This cmdlet gathers info about a computer(s) Hostname, IP Addresses and Speed, Mac Address, OS Version, Build Number, Service Pack Version, OS Architecture, Processor Architecture and Last Logged on User.

	.PARAMETER  ComputerName
		Enter a IP address, DNS name or Array of the machines.
    
    .PARAMETER Credential
        Use this Parameter to supply alternative creds that will work on a remote machine.

	.EXAMPLE
		PS C:\> Get-NTSystemInfo -ComputerName IAMACOMPUTER
        This gathers info about a specific machine "IAMCOMPUTER"
		
	.EXAMPLE
		PS C:\> Get-NTSystemInfo -ComputerName (Get-Content -Path C:\FileWithAComputerList.txt)

        This get the system info from a list of machines in the txt file "FileWithAComputerList.txt".
    
    .EXAMPLE
        PS c:\> Get-NTSystemInfo -ComputerName RemoteMachine -Credential (Get-Credential)

        This gathers info about a specific machine " RemoteMachine" Using alternate Credentials.

	.NOTES
		This funcion contains one parameter -ComputerName that can be pipped to.

#>

Function Get-NTSystemInfo {

[CmdletBinding()]

Param(

    [Parameter(Mandatory=$false,
     ValueFromPipeline=$True,
     HelpMessage="Enter Computer name or IP address to query")]
    [String[]] $ComputerName = 'localhost',
    
    [Parameter(ValueFromPipeline=$True)]
    [Object]$Credential

)
BEGIN{}

PROCESS
{

if ($Credential){
foreach ($Computer in $ComputerName) {
        $OSInfo = Get-WmiObject -class Win32_OperatingSystem -ComputerName $Computer -Credential $Credential
        $NetHardwareInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer -Credential $Credential
        $HardwareInfo = Get-WmiObject -Class Win32_Processor -ComputerName $Computer -Credential $Credential
        $NetHardwareInfo1 = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $Computer -Credential $Credential
        $UserSystemInfo = Get-WmiObject  -class win32_computerSystem -ComputerName $Computer -Credential $Credential

$Props =  @{'Hostname'=$OSInfo.CSName;
            'OS Version'=$OSInfo.name;
            'Build Number'=$OSInfo.BuildNumber;
            'Service Pack'=$OSInfo.ServicePackMajorVersion;
            'IP Addresses'=$NetHardwareInfo  | Sort-Object -Property IPAddress | Where-Object -Property IPAddress -GT $NULL | Select-Object -ExpandProperty IPAddress;
            'Mac Addresses'=$NetHardwareInfo | Sort-Object -Property MACAddress | Where-Object -Property MACAddress -GT $NULL | Select-Object -ExpandProperty MACAddress;
			'Network Speed'=$NetHardwareInfo1 | Where-Object -Property Speed -GT $NULL | Select-Object -ExpandProperty Speed ;
			'OS Architecture'=$OSInfo.OSArchitecture;
			'Processor Architecture'=$HardwareInfo.DataWidth;
			'Logged-in User' = $UserSystemInfo.Username;
            }
            $Object = New-Object -TypeName PSObject -Property $Props
Write-Output $Object | Select-Object -Property Hostname,'Logged-in User','IP Addresses','Network Speed','Mac Addresses','OS Version','Build Number','Service Pack','OS Architecture','Processor Architecture'


}} else {
foreach ($Computer in $ComputerName) {
        $OSInfo = Get-WmiObject -class Win32_OperatingSystem -ComputerName $Computer
        $NetHardwareInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer
        $HardwareInfo = Get-WmiObject -Class Win32_Processor -ComputerName $Computer
        $NetHardwareInfo1 = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $Computer
        $UserSystemInfo = Get-WmiObject  -class win32_computerSystem -ComputerName $Computer
        
$Props =  @{'Hostname'=$OSInfo.CSName;
            'OS Version'=$OSInfo.name;
            'Build Number'=$OSInfo.BuildNumber;
            'Service Pack'=$OSInfo.ServicePackMajorVersion;
            'IP Addresses'=$NetHardwareInfo  | Sort-Object -Property IPAddress | Where-Object -Property IPAddress -GT $NULL | Select-Object -ExpandProperty IPAddress;
            'Mac Addresses'=$NetHardwareInfo | Sort-Object -Property MACAddress | Where-Object -Property MACAddress -GT $NULL | Select-Object -ExpandProperty MACAddress;
			'Network Speed'=$NetHardwareInfo1 | Where-Object -Property Speed -GT $NULL | Select-Object -ExpandProperty Speed ;
			'OS Architecture'=$OSInfo.OSArchitecture;
			'Processor Architecture'=$HardwareInfo.DataWidth;
			'Logged-in User' = $UserSystemInfo.Username;
           }
           }
            
$Object = New-Object -TypeName PSObject -Property $Props
Write-Output $Object | Select-Object -Property Hostname,'Logged-in User','IP Addresses','Network Speed','Mac Addresses','OS Version','Build Number','Service Pack','OS Architecture','Processor Architecture'

}
}
}
Function Get-PendingReboot
{
<# 
.SYNOPSIS 
    Gets the pending reboot status on a local or remote computer. 
 
.DESCRIPTION 
    This Function will query the registry on a local or remote computer and determine if the 
    system is pending a reboot, from either Microsoft Patching or a Software Installation. 
    For Windows 2008+ the Function will query the CBS registry key as another factor in determining 
    pending reboot state.  "PendingFileRenameOperations" and "Auto Update\RebootRequired" are observed 
    as being consistant across Windows Server 2003 & 2008. 
   
    CBServicing = Component Based Servicing (Windows 2008) 
    WindowsUpdate = Windows Update / Auto Update (Windows 2003 / 2008) 
    CCMClientSDK = SCCM 2012 Clients only (DetermineIfRebootPending method) otherwise $null value 
    PendFileRename = PendingFileRenameOperations (Windows 2003 / 2008) 
 
.PARAMETER ComputerName 
    A single Computer or an array of computer names.  The default is localhost ($env:COMPUTERNAME). 
 
.PARAMETER ErrorLog 
    A single path to send error data to a log file. 
 
.EXAMPLE 
    PS C:\> Get-PendingReboot -ComputerName (Get-Content C:\ServerList.txt) | Format-Table -AutoSize 
   
    Computer CBServicing WindowsUpdate CCMClientSDK PendFileRename PendFileRenVal RebootPending 
    -------- ----------- ------------- ------------ -------------- -------------- ------------- 
    DC01     False   False           False      False 
    DC02     False   False           False      False 
    FS01     False   False           False      False 
 
    This example will capture the contents of C:\ServerList.txt and query the pending reboot 
    information from the systems contained in the file and display the output in a table. The 
    null values are by design, since these systems do not have the SCCM 2012 client installed, 
    nor was the PendingFileRenameOperations value populated. 
 
.EXAMPLE 
    PS C:\> Get-PendingReboot 
   
    Computer     : WKS01 
    CBServicing  : False 
    WindowsUpdate      : True 
    CCMClient    : False 
    PendComputerRename : False 
    PendFileRename     : False 
    PendFileRenVal     :  
    RebootPending      : True 
   
    This example will query the local machine for pending reboot information. 
   
.EXAMPLE 
    PS C:\> $Servers = Get-Content C:\Servers.txt 
    PS C:\> Get-PendingReboot -Computer $Servers | Export-Csv C:\PendingRebootReport.csv -NoTypeInformation 
   
    This example will create a report that contains pending reboot information. 
 
.LINK 
    Component-Based Servicing: 
    http://technet.microsoft.com/en-us/library/cc756291(v=WS.10).aspx 
   
    PendingFileRename/Auto Update: 
    http://support.microsoft.com/kb/2723674 
    http://technet.microsoft.com/en-us/library/cc960241.aspx 
    http://blogs.msdn.com/b/hansr/archive/2006/02/17/patchreboot.aspx 
 
    SCCM 2012/CCM_ClientSDK: 
    http://msdn.microsoft.com/en-us/library/jj902723.aspx 
 
.NOTES 
    Author:  Brian Wilhite 
    Email:   bcwilhite (at) live.com 
    Date:    29AUG2012 
    PSVer:   2.0/3.0/4.0/5.0 
    Updated: 01DEC2014 
    UpdNote: Added CCMClient property - Used with SCCM 2012 Clients only 
       Added ValueFromPipelineByPropertyName=$true to the ComputerName Parameter 
       Removed $Data variable from the PSObject - it is not needed 
       Bug with the way CCMClientSDK returned null value if it was false 
       Removed unneeded variables 
       Added PendFileRenVal - Contents of the PendingFileRenameOperations Reg Entry 
       Removed .Net Registry connection, replaced with WMI StdRegProv 
       Added ComputerPendingRename 
#>	
	
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[Alias("CN", "Computer")]
		[String[]]$ComputerName = "$env:COMPUTERNAME",
		
		[String]$ErrorLog
	)
	
	Begin { } ## End Begin Script Block 
	Process
	{
		Foreach ($Computer in $ComputerName)
		{
			Try
			{
				## Setting pending values to false to cut down on the number of else statements 
				$CompPendRen, $PendFileRename, $Pending, $SCCM = $false, $false, $false, $false
				
				## Setting CBSRebootPend to null since not all versions of Windows has this value 
				$CBSRebootPend = $null
				
				## Querying WMI for build version 
				$WMI_OS = Get-WmiObject -Class Win32_OperatingSystem -Property BuildNumber, CSName -ComputerName $Computer -ErrorAction Stop
				
				## Making registry connection to the local/remote computer 
				$HKLM = [UInt32] "0x80000002"
				$WMI_Reg = [WMIClass] "\\$Computer\root\default:StdRegProv"
				
				## If Vista/2008 & Above query the CBS Reg Key 
				If ([Int32]$WMI_OS.BuildNumber -ge 6001)
				{
					$RegSubKeysCBS = $WMI_Reg.EnumKey($HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\")
					$CBSRebootPend = $RegSubKeysCBS.sNames -contains "RebootPending"
				}
				
				## Query WUAU from the registry 
				$RegWUAURebootReq = $WMI_Reg.EnumKey($HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\")
				$WUAURebootReq = $RegWUAURebootReq.sNames -contains "RebootRequired"
				
				<## Query PendingFileRenameOperations from the registry 
				$RegSubKeySM = $WMI_Reg.GetMultiStringValue($HKLM, "SYSTEM\CurrentControlSet\Control\Session Manager\", "PendingFileRenameOperations")
				$RegValuePFRO = $RegSubKeySM.sValue#>
				
				## Query ComputerName and ActiveComputerName from the registry 
				$ActCompNm = $WMI_Reg.GetStringValue($HKLM, "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName\", "ComputerName")
				$CompNm = $WMI_Reg.GetStringValue($HKLM, "SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\", "ComputerName")
				If ($ActCompNm -ne $CompNm)
				{
					$CompPendRen = $true
				}
				
				## If PendingFileRenameOperations has a value set $RegValuePFRO variable to $true 
				If ($RegValuePFRO)
				{
					$PendFileRename = $true
				}
				
				## Determine SCCM 2012 Client Reboot Pending Status 
				## To avoid nested 'if' statements and unneeded WMI calls to determine if the CCM_ClientUtilities class exist, setting EA = 0 
				$CCMClientSDK = $null
				$CCMSplat = @{
					NameSpace = 'ROOT\ccm\ClientSDK'
					Class ='CCM_ClientUtilities'
					Name = 'DetermineIfRebootPending'
					ComputerName = $Computer
					ErrorAction = 'Stop'
				}
				## Try CCMClientSDK 
				Try
				{
					$CCMClientSDK = Invoke-WmiMethod @CCMSplat
				}
				Catch [System.UnauthorizedAccessException] {
					$CcmStatus = Get-Service -Name CcmExec -ComputerName $Computer -ErrorAction SilentlyContinue
					If ($CcmStatus.Status -ne 'Running')
					{
						Write-Warning "$Computer`: Error - CcmExec service is not running."
						$CCMClientSDK = $null
					}
				}
				Catch
				{
					$CCMClientSDK = $null
				}
				
				If ($CCMClientSDK)
				{
					If ($CCMClientSDK.ReturnValue -ne 0)
					{
						Write-Warning "Error: DetermineIfRebootPending returned error code $($CCMClientSDK.ReturnValue)"
					}
					If ($CCMClientSDK.IsHardRebootPending -or $CCMClientSDK.RebootPending)
					{
						$SCCM = $true
					}
				}
				
				Else
				{
					$SCCM = $null
				}
				
				## Creating Custom PSObject and Select-Object Splat 
				$SelectSplat = @{
					Property = (
					'Computer',
					'CBServicing',
					'WindowsUpdate',
					'CCMClientSDK',
					'PendComputerRename',
					'PendFileRename',
					'PendFileRenVal',
					'RebootPending'
					)
				}
				New-Object -TypeName PSObject -Property @{
					Computer = $WMI_OS.CSName
					CBServicing = $CBSRebootPend
					WindowsUpdate = $WUAURebootReq
					CCMClientSDK = $SCCM
					PendComputerRename = $CompPendRen
					PendFileRename = $PendFileRename
					PendFileRenVal = $RegValuePFRO
					RebootPending = ($CompPendRen -or $CBSRebootPend -or $WUAURebootReq -or $SCCM -or $PendFileRename)
				} | Select-Object @SelectSplat
				
			}
			Catch
			{
				Write-Warning "$Computer`: $_"
				## If $ErrorLog, log the file to a user specified location/path 
				If ($ErrorLog)
				{
					Out-File -InputObject "$Computer`,$_" -FilePath $ErrorLog -Append
				}
			}
		} ## End Foreach ($Computer in $ComputerName)       
	} ## End Process 
	
	End { } ## End End 
	
} ## End Function Get-PendingReboot
Function Get-PSCredential{

    [CmdletBinding()]
    param(
    [parameter(
        position = 0,
        mandatory = 0)]
    $credentialpath = "C:\Deployment\Bin\credential.json"
    )

    $credential = Get-Content $credentialpath -Raw | ConvertFrom-Json
    $secpasswd = ConvertTo-SecureString $credential.password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($($credential.user), $secpasswd)    

    return $credential
}


Function New-PSCredential{

    [CmdletBinding()]
    param(
    [parameter(
        position = 0,
        mandatory = 0)]
    $credentialpath = "C:\Deployment\Bin\credential.json",

    [parameter(
        position = 1,
        mandatory)]
    $user,

    [parameter(
        position = 2,
        mandatory)]
    $password
    )

    [PSCustomObject]@{
        user = $user
        password = $password
    } | ConvertTo-Json | Out-File -FilePath $credentialpath -Force

}


New-PSCredential -user ec2-user -password Qwe09r7c23
Get-PSCredential
Function Get-PrivilegedGroupsMemberCount 
{
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
Function Get-ProcessForeignAddress
{
<#
.SYNOPSIS
	Get all foreignIPAddress for all or specific processname
	
.DESCRIPTION
	Get all foreignIPAddress for all or specific processname
	
.PARAMETER ProcessName
	Specifies the ProcessName to filter on
	
.EXAMPLE
	Get-ProcessForeignAddress
	
	Retrieve all the foreign addresses
	
.EXAMPLE
	Get-ProcessForeignAddress chrome
	
	Show all the foreign address(es) for the process chrome
	
.EXAMPLE
	Get-ProcessForeignAddress chrome | select ForeignAddress -Unique
	
	Show all the foreign address(es) for the process chrome and show only the ForeignAddress(es) once
	
.NOTES
	Author	: Francois-Xavier Cat
	Website	: www.lazywinadmin.com
	Github	: github.com/lazywinadmin
	Twitter	: @lazywinadm
#>
	PARAM ($ProcessName)
	$netstat = netstat -no
	
	$Result = $netstat[4..$netstat.count] |
	ForEach-Object {
		$current = $_.trim() -split '\s+'
		
		New-Object -TypeName PSobject -Property @{
			ProcessName = (Get-Process -id $current[4]).processname
			ForeignAddressIP = ($current[2] -split ":")[0] #-as [ipaddress]
			ForeignAddressPort = ($current[2] -split ":")[1]
			State = $current[3]
		}
	}
	
	if ($ProcessName)
	{
		$result | Where-Object { $_.processname -like "$processname" }
	}
	else { $Result }
}
Function Get-ProductId
{
    [CmdletBinding()]
    param
    (
        [parameter(
            Mandatory = 0,
            Position  = 1,
            ValueFromPipeline = 1,
            ValueFromPipelineByPropertyName =1)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $path = "registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    begin
    {
        $list = New-Object 'System.Collections.Generic.List[PSCustomObject]'
    }

    process
    {
        $path `
        | %{
            # validation
            if (!(Test-Path $_))
            {
                throw "path '{0}' not found Exception!!" -f $_
            }

            $reg = Get-ItemProperty -Path $_ | where DisplayName
            $reg `
            | %{
                $obj = [PSCustomObject]@{
                    DisplayName    = $_.DisplayName
                    DisplayVersion = $_.DisplayVersion
                    Publisher      = $_.Publisher
                    InstallDate    = $_ | where {$_.InstallDate} | %{[DateTime]::ParseExact($_.InstallDate,"yyyyMMdd",$null)}
                    ProductId      = $_.PSChildName | %{$_ -replace "{" -replace "}"}
                }
                $list.Add($obj)
            }
        }
    }

    end
    {
        $list | sort DisplayName
    }
}
Function Get-PSObjectEmptyOrNullProperty
{
<#
.SYNOPSIS
	Function to Get all the empty or null properties with empty value in a PowerShell Object
	
.DESCRIPTION
	Function to Get all the empty or null properties with empty value in a PowerShell Object.
	I used this Function in a System Center Orchestrator where I had a runbook that could update most of the important 
	properties of a user. Using this Function I knew which properties were not be updated.
	
.PARAMETER PSObject
	Specifies the PowerShell Object
	
.EXAMPLE
	PS C:\> Get-PSObjectEmptyOrNullProperty -PSObject $UserInfo
	
.EXAMPLE

    # Create a PowerShell Object with some properties
    $o=''|select FirstName,LastName,nullable
    $o.firstname='Nom'
    $o.lastname=''
    $o.nullable=$null
	
    # Look for empty or null properties
    Get-PSObjectEmptyOrNullProperty -PSObject $o

    MemberType      : NoteProperty
    IsSettable      : True
    IsGettable      : True
    Value           : 
    TypeNameOfValue : System.String
    Name            : LastName
    IsInstance      : True

    MemberType      : NoteProperty
    IsSettable      : True
    IsGettable      : True
    Value           : 
    TypeNameOfValue : System.Object
    Name            : nullable
    IsInstance      : True

.NOTES
	Francois-Xavier Cat	
	www.lazywinadmin.com
	@lazywinadm
#>
	PARAM (
		$PSObject)
	PROCESS
	{
		$PsObject.psobject.Properties |
		Where-Object { -not $_.value }
	}
}
Function Get-RegistryKeyPropertiesAndValues
{
  <#
    This Function is used here to retrieve registry values while omitting the PS properties
    Example: Get-RegistryKeyPropertiesAndValues -path 'HKCU:\Volatile Environment'
    Origin: Http://www.ScriptingGuys.com/blog
    Via: http://stackoverflow.com/questions/13350577/can-powershell-get-childproperty-get-a-list-of-real-registry-keys-like-reg-query
  #>

 Param(
  [Parameter(Mandatory=$true)]
  [string]$path
  )

  Push-Location
  Set-Location -Path $path
  Get-Item . |
  Select-Object -ExpandProperty property |
  ForEach-Object {
      New-Object psobject -Property @{"Folder"=$_;
        "RedirectedLocation" = (Get-ItemProperty -Path . -Name $_).$_}}
  Pop-Location
}

# Get the user profile path, while escaping special characters because we are going to use the -match operator on it
$Profilepath = [regex]::Escape($env:USERPROFILE)

# List all folders
$RedirectedFolders = Get-RegistryKeyPropertiesAndValues -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" | Where-Object {$_.RedirectedLocation -notmatch "$Profilepath"}
if ($RedirectedFolders -eq $null) {
    Write-Output "No folders are redirected for this user"
} else {
    $RedirectedFolders | format-list *
}
###########################################################
#Script Title: Get Registry Key and Values PowerShell Tool
#Script File Name: Get-RegKeyandValues.ps1
#Author: Ron Ratzlaff (aka "The_Ratzenator")
#Date Created: 6/22/2014
###########################################################

#Requires -Version 3.0

Function Get-RegKeyandValues
{
   <#
	  .SYNOPSIS
	  
	  	The "Get-RegKeyandValues" Function will attempt to retrieve Registry keys and values that you specify, if they exist 
	  
	  .EXAMPLE

		To get the values and the sub keys within the "MyRegKey" key on the local computer, use the following syntax:

		Get-RegKeyandValues -RegHive "HKLM" -RegKey "SOFTWARE\MyRegKey" -GetRegKeyVals "Yes" -GetRegSubKeys "Yes"

	  .EXAMPLE
		
		To get the values, but exclude the sub keys within the "MyRegKey" key on the local computer, use the following syntax:
		
		Get-RegKeyandValues -RegHive "HKLM" -RegKey "SOFTWARE\MyRegKey" -GetRegKeyVals "Yes" -GetRegSubKeys "No"

	  .EXAMPLE

		To get the sub keys, but exclude the values within the "MyRegKey" key on the local computer, use the following syntax:
		
		Get-RegKeyandValues -ComputerName "Computer1" -RegHive "HKLM" -RegKey "SOFTWARE\MyRegKey" -GetRegKeyVals "No" -GetRegSubKeys "Yes"
	
	  .EXAMPLE
		
		To retrieve the Registry info remotely on more than one computer, you can use an array as shown in the following syntax:
		
		Get-RegKeyandValues -ComputerName @("Computer1", "Computer2") -RegHive "HKLM" -RegKey "SOFTWARE\Wow6432Node" -GetRegKeyVals "Yes" -GetRegSubKeys "Yes"

	  .EXAMPLE
		
		To retrieve the Registry info remotely on more than one computer, you can use a file (.csv, .txt) and use the Get-Content cmdlet as shown in the following syntax:
		
		Get-RegKeyandValues -ComputerName (Get-Content -Path "$env:TEMP\ComputerList.txt") -RegHive "HKLM" -RegKey "SOFTWARE\Wow6432Node" -GetRegKeyVals "Yes" -GetRegSubKeys "Yes"

	  .PARAMETER ComputerName
	  
	  	A mandatory parameter used to query a single computer or multiple computers.
	  
	  .PARAMETER RegHive
	  
	  	A mandatory parameter used to query one of the primary Registry Hives (HKCR, HKCU, HKLM, HKU, and HKCC). You must specify one of the primary Registry Hives as shown below, or an error will display:
		
		HKCR: HKEY_CLASSES_ROOT
		HKCU: HKEY_CURRENT_USER
		HKLM: HKEY_LOCAL_MACHINE
		HKU: HKEY_USERS
		HKCC: HKEY_CURRENT_CONFIG
	  
	  .PARAMETER RegKey
	  
	  	A manadatory parameter used to query a specified key. A path must be specified for the key. For instance, to query the "Uninstall" key located under HKLM, you will need to specify the following path:
		
			"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	  
	  .PARAMETER GetRegKeyVals
	  
	  	An optional parameter (not mandatory) that retrieves the values listed within the specified Registry Key parameter (-RegKey) that when used, requires that either a "Yes" or a "No" value is specified
	  
	  .PARAMETER GetRegSubKeys
	  	
		An optional parameter (not mandatory) that retrieves the sub keys listed within the specified Registry Key parameter (-RegKey) that when used, requires that either a "Yes" or a "No" value is specified
  #>
   
   [cmdletbinding()]            
	
	Param
	(            
	 	[Parameter(Position=0,
		 ValueFromPipeline,
	  	 ValueFromPipelineByPropertyName,
	  	 HelpMessage='What computer name would you like to target?')]
	  	$ComputerName = $env:COMPUTERNAME,
		
		[Parameter(Mandatory=$true,
	  	 Position=1,
		 ValueFromPipeline,
	  	 ValueFromPipelineByPropertyName,
	  	 HelpMessage='What Registry Hive would you like to target?')]
		 [ValidateSet('HKCR','HKCU','HKLM','HKU',IgnoreCase = $true)]
		 [ValidateNotNullOrEmpty()]
		[string[]]$RegHive,
		
		[Parameter(Mandatory=$true,
		 Position=2,
		 ValueFromPipeline,
	  	 ValueFromPipelineByPropertyName,
	  	 HelpMessage='What Registry Key would you like to target?')]
		 [ValidateNotNullOrEmpty()] 
		[string[]]$RegKey,
		
		[Parameter(Position=3,
	  	 HelpMessage='Would you like to display the Registry Values within the "$RegKey" Key?')]
		[ValidateSet('Yes', 'No', IgnoreCase = $true)]
		[string[]]$GetRegKeyVals,
		
		[Parameter(Position=4,
	  	 HelpMessage='Would you like to display the Registry Sub Keys under the "$RegKey" Key?')]
		[ValidateSet('Yes', 'No',IgnoreCase = $true)]
		[string[]]$GetRegSubKeys
	)   
   
    Begin {}
   
    Process
	{
		$NewLine = "`r`n"
		
		Switch ($RegHive)
		{
			"HKCR"
			{
				$Hive = "ClassesRoot"
			}
			
			"HKCU"
			{
				$Hive = "CurrentUsers"
			}
			
			"HKLM"
			{
				$Hive = "LocalMachine"
			}
			
			"HKU"
			{
				$Hive = "Users"
			}
			
			"HKCC"
			{
				$Hive = "CurrentConfig"
			}
		}
		
		Foreach ($Computer in $ComputerName)
		{
			$RegHiveType = [Microsoft.Win32.RegistryHive]::$Hive
			$OpenBaseRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHiveType, $Computer)
			$OpenRegSubKey = $OpenBaseRegKey.OpenSubKey($RegKey)
			
			If (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
			{
				If ($OpenRegSubKey)
				{			
					If ($GetRegKeyVals -match "Yes" -and $GetRegSubKeys -match "Yes")
					{
						$NewLine 
					
						Write-Output "Computer"
						Write-Output "--------"
						
						$NewLine
						
						Write-Output "$Computer"
						
						$NewLine
						
						Write-Output "Reg Key Exist"
						Write-Output "-------------"
						
						$NewLine
						
						Write-Output "'$RegKey'"
						
						$NewLine
						
						Write-Output "Reg Key Values"
						Write-Output "--------------"
						
						$NewLine
						
						$GetRegKeyVal = Foreach($RegKeyValue in $OpenRegSubKey.GetValueNames()){$RegKeyValue}
						
						If ($GetRegKeyVal -ne $null)
						{
							$GetRegKeyVal
						}
						
						Else
						{
							Write-Output "No Registry Values exist within the $RegKey"
						}
						
						$NewLine
						
						Write-Output "Reg Sub Keys"
						Write-Output "------------"
						
						$NewLine
						
						$GetRegSubKey = Foreach($RegSubKey in $OpenRegSubKey.GetSubKeyNames()){$RegSubKey}
						
						If ($GetRegSubKey -ne $null)
						{
							$GetRegSubKey
						}
						
						Else
						{
							Write-Output "No Registry Sub Keys exist under the $RegKey"
						}
						
						$NewLine
					}
					
					ElseIf ($GetRegKeyVals -match "Yes" -and $GetRegSubKeys -match "No")
					{
						$NewLine 
					
						Write-Output "Computer"
						Write-Output "--------"
						
						$NewLine
						
						Write-Output "$Computer"
						
						$NewLine
						
						Write-Output "Reg Key Exist"
						Write-Output "-------------"
						
						$NewLine
						
						Write-Output "'$RegKey'"
						
						$NewLine
						
						Write-Output "Reg Key Values"
						Write-Output "--------------"
						
						$NewLine
						
						$GetRegKeyVal = Foreach($RegKeyValue in $OpenRegSubKey.GetValueNames()){$RegKeyValue}
						
						If ($GetRegKeyVal -ne $null)
						{
							$GetRegKeyVal
						}
						
						Else
						{
							Write-Output "No Registry Values exist within the $RegKey"
						}
						
						$NewLine
					}
					
					ElseIf ($GetRegKeyVals -match "No" -and $GetRegSubKeys -match "Yes")
					{
						$NewLine 
					
						Write-Output "Computer"
						Write-Output "--------"
						
						$NewLine
						
						Write-Output "$Computer"
						
						$NewLine
						
						Write-Output "Reg Key Exist"
						Write-Output "-------------"
						
						$NewLine
						
						Write-Output "'$RegKey'"
						
						$NewLine
						
						Write-Output "Reg Sub Keys"
						Write-Output "------------"
						
						$NewLine
						
						$GetRegSubKey = Foreach($RegSubKey in $OpenRegSubKey.GetSubKeyNames()){$RegSubKey}
						
						If ($GetRegSubKey -ne $null)
						{
							$GetRegSubKey
						}
						
						Else
						{
							Write-Output "No Registry Sub Keys exist under the $RegKey"
						}
						
						$NewLine
					}
					
					Else
					{
						$NewLine
						
						Write-Output "Computer"
						Write-Output "--------"
						
						$NewLine
						
						Write-Output "$Computer"
						
						$NewLine
						
						Write-Output "Reg Key Exist"
						Write-Output "-------------"
						
						$NewLine
						
						Write-Output "'$RegKey'"
						
						$NewLine
					}
				}
				
				Else
				{
					$NewLine
					
					Write-Output "Computer"
					Write-Output "--------"
					
					$NewLine
					
					Write-Output "$Computer"
					
					$NewLine
					
					Write-Warning -Message "Could not find $RegKey"
				
					$NewLine
				}
			}
			
			Else
			{
				$NewLine
				
				Write-Warning -Message "$Computer is offline!"
				
				$NewLine
			}
		}
	}
	
	End {}
}#EndFunction
	
Function Get-RemoteAppliedGPOs
{
    <#
    .SYNOPSIS
       Gather applied GPO information from local or remote systems.
    .DESCRIPTION
       Gather applied GPO information from local or remote systems. Can utilize multiple runspaces and 
       alternate credentials.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information

    .EXAMPLE
       $a = Get-RemoteAppliedGPOs
       $a.AppliedGPOs | 
            Select Name,AppliedOrder |
            Sort-Object AppliedOrder
       
       Name                            appliedOrder
       ----                            ------------
       Local Group Policy                         1
       
       Description
       -----------
       Get all the locally applied GPO information then display them in their applied order.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 09/01/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
       
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of Function")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the Function to prompt for alternate credentials")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    BEGIN
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Remote Applied GPOs: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Remote Applied GPOs: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Remote Applied GPOs: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Remote Applied GPOs: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Remote Applied GPOs: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Remote Applied GPOs: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,
 
                [Parameter(Position=1)]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }

                # General variables
                $GPOPolicies = @()
                $PSDateTime = Get-Date
                
                #region GPO Data

                $GPOQuery = Get-WmiObject @WMIHast `
                                          -Namespace "ROOT\RSOP\Computer" `
                                          -Class RSOP_GPLink `
                                          -Filter "AppliedOrder <> 0" |
                            Select @{n='linkOrder';e={$_.linkOrder}},
                                   @{n='appliedOrder';e={$_.appliedOrder}},
                                   @{n='GPO';e={$_.GPO.ToString().Replace("RSOP_GPO.","")}},
                                   @{n='Enabled';e={$_.Enabled}},
                                   @{n='noOverride';e={$_.noOverride}},
                                   @{n='SOM';e={[regex]::match( $_.SOM , '(?<=")(.+)(?=")' ).value}},
                                   @{n='somOrder';e={$_.somOrder}}
                foreach($GP in $GPOQuery)
                {
                    $AppliedPolicy = Get-WmiObject @WMIHast `
                                                   -Namespace 'ROOT\RSOP\Computer' `
                                                   -Class 'RSOP_GPO' -Filter $GP.GPO
                        $ObjectProp = @{
                            'Name' = $AppliedPolicy.Name
                            'GuidName' = $AppliedPolicy.GuidName
                            'ID' = $AppliedPolicy.ID
                            'linkOrder' = $GP.linkOrder
                            'appliedOrder' = $GP.appliedOrder
                            'Enabled' = $GP.Enabled
                            'noOverride' = $GP.noOverride
                            'SourceOU' = $GP.SOM
                            'somOrder' = $GP.somOrder
                        }
                        
                        $GPOPolicies += New-Object PSObject -Property $ObjectProp
                }
                          
                Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: Share session information' -f $ComputerName)

                # Modify this variable to change your default set of display properties
                $defaultProperties    = @('ComputerName','AppliedGPOs')
                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'AppliedGPOs' = $GPOPolicies
                }
                $ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
                
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.AppliedGPOs.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                #endregion GPO Data

                Write-Output -InputObject $ResultObject
            }
            catch
            {
                Write-Warning -Message ('Remote Applied GPOs: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: End' -f $ComputerName)
        }
 
        Function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers.($runspace.ID)
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Remote Applied GPOs: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Remote Applied GPOs: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Remote Applied GPOs: Getting info'
                        Status = 'Remote Applied GPOs: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    PROCESS
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Remote Applied GPOs: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
    END
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Remote Applied GPOs: Getting share session information' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Remote Applied GPOs: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}
<#
.Synopsis
   Gets Disk Space of the given remote computer name
.DESCRIPTION
   Get-RemoteComputerDisk cmdlet gets the used, free and total space with the drive name.
.EXAMPLE
   Get-RemoteComputerDisk -RemoteComputerName "abc.contoso.com"
   Drive    UsedSpace(in GB)    FreeSpace(in GB)    TotalSpace(in GB)
   C        75                  52                  127
   D        28                  372                 400

.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FunctionALITY
   The Functionality that best describes this cmdlet
#>
Function Get-RemoteComputerDisk
{
    
    Param
    (
        $RemoteComputerName
    )

    Begin
    {
        $output="Drive `t UsedSpace(in GB) `t FreeSpace(in GB) `t TotalSpace(in GB) `n"
    }
    Process
    {
        $drives=Get-WmiObject Win32_LogicalDisk -ComputerName $RemoteComputerName

        foreach ($drive in $drives){
            
            $drivename=$drive.DeviceID
            $freespace=[int]($drive.FreeSpace/1GB)
            $totalspace=[int]($drive.Size/1GB)
            $usedspace=$totalspace - $freespace
            $output=$output+$drivename+"`t`t"+$usedspace+"`t`t`t`t`t`t"+$freespace+"`t`t`t`t`t`t"+$totalspace+"`n"
        }
    }
    End
    {
        return $output
    }
}
Function Get-RemoteProgram {
<#
.Synopsis
Generates a list of installed programs on a computer

.DESCRIPTION
This Function generates a list by querying the registry and returning the installed programs of a local or remote computer.

.NOTES   
Name       : Get-RemoteProgram
Author     : Jaap Brasser
Version    : 1.3
DateCreated: 2013-08-23
DateUpdated: 2016-08-26
Blog       : http://www.jaapbrasser.com

.LINK
http://www.jaapbrasser.com

.PARAMETER ComputerName
The computer to which connectivity will be checked

.PARAMETER Property
Additional values to be loaded from the registry. Can contain a string or an array of string that will be attempted to retrieve from the registry for each program entry

.PARAMETER ExcludeSimilar
This will filter out similar programnames, the default value is to filter on the first 3 words in a program name. If a program only consists of less words it is excluded and it will not be filtered. For example if you Visual Studio 2015 installed it will list all the components individually, using -ExcludeSimilar will only display the first entry.

.PARAMETER SimilarWord
This parameter only works when ExcludeSimilar is specified, it changes the default of first 3 words to any desired value.

.EXAMPLE
Get-RemoteProgram

Description:
Will generate a list of installed programs on local machine

.EXAMPLE
Get-RemoteProgram -ComputerName server01,server02

Description:
Will generate a list of installed programs on server01 and server02

.EXAMPLE
Get-RemoteProgram -ComputerName Server01 -Property DisplayVersion,VersionMajor

Description:
Will gather the list of programs from Server01 and attempts to retrieve the displayversion and versionmajor subkeys from the registry for each installed program

.EXAMPLE
'server01','server02' | Get-RemoteProgram -Property Uninstallstring

Description
Will retrieve the installed programs on server01/02 that are passed on to the Function through the pipeline and also retrieves the uninstall string for each program

.EXAMPLE
'server01','server02' | Get-RemoteProgram -Property Uninstallstring -ExcludeSimilar -SimilarWord 4

Description
Will retrieve the installed programs on server01/02 that are passed on to the Function through the pipeline and also retrieves the uninstall string for each program. Will only display a single entry of a program of which the first four words are identical.
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline              =$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0
        )]
        [string[]]
            $ComputerName = $env:COMPUTERNAME,
        [Parameter(Position=0)]
        [string[]]
            $Property,
        [switch]
            $ExcludeSimilar,
        [int]
            $SimilarWord
    )

    begin {
        $RegistryLocation = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
                            'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        $HashProperty = @{}
        $SelectProperty = @('ProgramName','ComputerName')
        if ($Property) {
            $SelectProperty += $Property
        }
    }

    process {
        foreach ($Computer in $ComputerName) {
            $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
            $RegistryLocation | ForEach-Object {
                $CurrentReg = $_
                if ($RegBase) {
                    $CurrentRegKey = $RegBase.OpenSubKey($CurrentReg)
                    if ($CurrentRegKey) {
                        $CurrentRegKey.GetSubKeyNames() | ForEach-Object {
                            if ($Property) {
                                foreach ($CurrentProperty in $Property) {
                                    $HashProperty.$CurrentProperty = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue($CurrentProperty)
                                }
                            }
                            $HashProperty.ComputerName = $Computer
                            $HashProperty.ProgramName = ($DisplayName = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue('DisplayName'))
                            if ($DisplayName) {
                                New-Object -TypeName PSCustomObject -Property $HashProperty |
                                Select-Object -Property $SelectProperty
                            } 
                        }
                    }
                }
            } | ForEach-Object -Begin {
                if ($SimilarWord) {
                    $Regex = [regex]"(^(.+?\s){$SimilarWord}).*$|(.*)"
                } else {
                    $Regex = [regex]"(^(.+?\s){3}).*$|(.*)"
                }
                [System.Collections.ArrayList]$Array = @()
            } -Process {
                if ($ExcludeSimilar) {
                    $null = $Array.Add($_)
                } else {
                    $_
                }
            } -End {
                if ($ExcludeSimilar) {
                    $Array | Select-Object -Property *,@{
                        name       = 'GroupedName'
                        expression = {
                            ($_.ProgramName -split $Regex)[1]
                        }
                    } |
                    Group-Object -Property 'GroupedName' | ForEach-Object {
                        $_.Group[0] | Select-Object -Property * -ExcludeProperty GroupedName
                    }
                }
            }
        }
    }
}
## Get-RemoteRegistry
########################################################################################
## Version: 2.1
##  + Fixed a pasting bug 
##  + I added the "Properties" parameter so you can select specific registry values
## NOTE: you have to have access, and the remote registry service has to be running
########################################################################################
## USAGE:
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP"
##     * Returns a list of subkeys (because this key has no properties)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727"
##     * Returns a list of subkeys and all the other "properties" of the key
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\Version"
##     * Returns JUST the full version of the .Net SP2 as a STRING (to preserve prior behavior)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" Version
##     * Returns a custom object with the property "Version" = "2.0.50727.3053" (your version)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" Version,SP
##     * Returns a custom object with "Version" and "SP" (Service Pack) properties
##
##  For fun, get all .Net Framework versions (2.0 and greater) 
##  and return version + service pack with this one command line:
##
##    Get-RemoteRegistry $RemotePC "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP" | 
##    Select -Expand Subkeys | ForEach-Object { 
##      Get-RemoteRegistry $RemotePC "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\$_" Version,SP 
##    }
##
########################################################################################
Function Get-RemoteRegistry {
param(
    [string]$computer = $(Read-Host "Remote Computer Name")
   ,[string]$Path     = $(Read-Host "Remote Registry Path (must start with HKLM,HKCU,etc)")
   ,[string[]]$Properties
   ,[switch]$Verbose
)
if ($Verbose) { $VerbosePreference = 2 } # Only affects this script.

   $root, $last = $Path.Split("\")
   $last = $last[-1]
   $Path = $Path.Substring($root.Length + 1,$Path.Length - ( $last.Length + $root.Length + 2))
   $root = $root.TrimEnd(":")

   #split the path to get a list of subkeys that we will need to access
   # ClassesRoot, CurrentUser, LocalMachine, Users, PerformanceData, CurrentConfig, DynData
   switch($root) {
      "HKCR"  { $root = "ClassesRoot"}
      "HKCU"  { $root = "CurrentUser" }
      "HKLM"  { $root = "LocalMachine" }
      "HKU"   { $root = "Users" }
      "HKPD"  { $root = "PerformanceData"}
      "HKCC"  { $root = "CurrentConfig"}
      "HKDD"  { $root = "DynData"}
      default { return "Path argument is not valid" }
   }


   #Access Remote Registry Key using the static OpenRemoteBaseKey method.
   Write-Verbose "Accessing $root from $computer"
   $rootkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($root,$computer)
   if(-not $rootkey) { Write-Error "Can't open the remote $root registry hive" }

   Write-Verbose "Opening $Path"
   $key = $rootkey.OpenSubKey( $Path )
   if(-not $key) { Write-Error "Can't open $($root + '\' + $Path) on $computer" }

   $subkey = $key.OpenSubKey( $last )
   
   $output = new-object object

   if($subkey -and $Properties -and $Properties.Count) {
      foreach($property in $Properties) {
         Add-Member -InputObject $output -Type NoteProperty -Name $property -Value $subkey.GetValue($property)
      }
      Write-Output $output
   } elseif($subkey) {
      Add-Member -InputObject $output -Type NoteProperty -Name "Subkeys" -Value @($subkey.GetSubKeyNames())
      foreach($property in $subkey.GetValueNames()) {
         Add-Member -InputObject $output -Type NoteProperty -Name $property -Value $subkey.GetValue($property)
      }
      Write-Output $output
   }
   else
   {
      $key.GetValue($last)
   }
 }
Function Get-RemoteRegistryKey {
    <#
    .SYNOPSIS
        Set registry key on remote computers.
    .DESCRIPTION
        This Function uses .Net class [Microsoft.Win32.RegistryKey].
    .PARAMETER ComputerName
        Name of the remote computers.
    .PARAMETER Hive
        Hive where the key is.
    .PARAMETER KeyPath
        Path of the key.
    .PARAMETER Name
        Name of the key setting.
    .PARAMETER Type
        Type of the key setting.
    .PARAMETER Value
        Value tu put in the key setting.
    .EXAMPLE
        Get-RemoteRegistryKey -ComputerName $env:ComputerName -Hive "LocalMachine" -KeyPath "software\corporate\master\Test" -Name "TestName" -Type String -Value "TestValue" -Verbose
    .LINK
        http://itfordummies.net
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$false,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$True,
            HelpMessage='Provide a ComputerName')]
        [String[]]$ComputerName=$env:ComputerName,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryHive]$Hive,
        
        [Parameter(Mandatory=$true)]
        [String]$KeyPath,
        
        [Parameter(Mandatory=$true)]
        [String]$Name,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryValueKind]$Type,
        
        [Parameter(Mandatory=$true)]
        [Object]$Value
    )
    Begin{
    }
    Process{
        ForEach ($Computer in $ComputerName) {
            try {
                Write-Verbose "Trying computer $Computer"
                $reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("$hive", "$Computer")
                Write-Debug -Message "Contenur de Reg $reg"
                $key=$reg.OpenSubKey("$KeyPath",$true)
                if($key -eq $null){
                    Write-Verbose -Message "Key not found."
                    Write-Verbose -Message "Calculating parent and child paths..."
                    $parent = Split-Path -Path $KeyPath -Parent
                    $child = Split-Path -Path $KeyPath -Leaf
                    Write-Verbose -Message "Creating the subkey $child in $parent..."
                    $Key=$reg.OpenSubKey("$parent",$true)
                    $Key.CreateSubKey("$child") | Out-Null
                    Write-Verbose -Message "Setting $value in $KeyPath"
                    $key=$reg.OpenSubKey("$KeyPath",$true)
                    $key.SetValue($Name,$Value,$Type)
                }
                else{
                    Write-Verbose "Key found, setting $Value in $KeyPath..."
                    $key.SetValue($Name,$Value,$Type)
                }
                Write-Verbose "$Computer done."
            }#End Try
            catch {Write-Warning "$Computer : $_"} 
        }#End ForEach
    }#End Process
    End{
    }
}
Function Get-ScheduledTask {   
    [cmdletbinding()]
    Param (    
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string[]]$Computername = $env:COMPUTERNAME
    )
    Begin {
        $ST = New-Object -com("Schedule.Service")
    }
    Process {
        ForEach ($Computer in $Computername) {
            Try {
                $st.Connect($Computer)
                $root=  $st.GetFolder("\")
                @($root.GetTasks(0)) | ForEach {
                    $xml = ([xml]$_.xml).task
                    [pscustomobject] @{
                        Computername = $Computer
                        Task = $_.Name
                        Author = $xml.RegistrationInfo.Author
                        RunAs = $xml.Principals.Principal.UserId                        
                        Enabled = $_.Enabled
                        State = Switch ($_.State) {
                            0 {'Unknown'}
                            1 {'Disabled'}
                            2 {'Queued'}
                            3 {'Ready'}
                            4 {'Running'}
                        }
                        LastTaskResult = Switch ($_.LastTaskResult) {
                            0x0 {"Successfully completed"}
                            0x1 {"Incorrect Function called"}
                            0x2 {"File not found"}
                            0xa {"Environment is not correct"}
                            0x41300 {"Task is ready to run at its next scheduled time"}
                            0x41301 {"Task is currently running"}
                            0x41302 {"Task is disabled"}
                            0x41303 {"Task has not yet run"}
                            0x41304 {"There are no more runs scheduled for this task"}
                            0x41306 {"Task is terminated"}
                            0x00041307 {"Either the task has no triggers or the existing triggers are disabled or not set"}
                            0x00041308 {"Event triggers do not have set run times"}
                            0x80041309 {"A task's trigger is not found"}
                            0x8004130A {"One or more of the properties required to run this task have not been set"}
                            0x8004130B {"There is no running instance of the task"}
                            0x8004130C {"The Task * SCHEDuler service is not installed on this computer"}
                            0x8004130D {"The task object could not be opened"}
                            0x8004130E {"The object is either an invalid task object or is not a task object"}
                            0x8004130F {"No account information could be found in the Task * SCHEDuler security database for the task indicated"}
                            0x80041310 {"Unable to establish existence of the account specified"}
                            0x80041311 {"Corruption was detected in the Task * SCHEDuler security database"}
                            0x80041312 {"Task * SCHEDuler security services are available only on Windows NT"}
                            0x80041313 {"The task object version is either unsupported or invalid"}
                            0x80041314 {"The task has been configured with an unsupported combination of account settings and run time options"}
                            0x80041315 {"The Task * SCHEDuler Service is not running"}
                            0x80041316 {"The task XML contains an unexpected node"}
                            0x80041317 {"The task XML contains an element or attribute from an unexpected namespace"}
                            0x80041318 {"The task XML contains a value which is incorrectly formatted or out of range"}
                            0x80041319 {"The task XML is missing a required element or attribute"}
                            0x8004131A {"The task XML is malformed"}
                            0x0004131B {"The task is registered, but not all specified triggers will start the task"}
                            0x0004131C {"The task is registered, but may fail to start"}
                            0x8004131D {"The task XML contains too many nodes of the same type"}
                            0x8004131E {"The task cannot be started after the trigger end boundary"}
                            0x8004131F {"An instance of this task is already running"}
                            0x80041320 {"The task will not run because the user is not logged on"}
                            0x80041321 {"The task image is corrupt or has been tampered with"}
                            0x80041322 {"The Task * SCHEDuler service is not available"}
                            0x80041323 {"The Task * SCHEDuler service is too busy to handle your request"}
                            0x80041324 {"The Task * SCHEDuler service attempted to run the task, but the task did not run due to one of the constraints in the task definition"}
                            0x00041325 {"The Task * SCHEDuler service has asked the task to run"}
                            0x80041326 {"The task is disabled"}
                            0x80041327 {"The task has properties that are not compatible with earlier versions of Windows"}
                            0x80041328 {"The task settings do not allow the task to start on demand"}
                            Default {[string]$_}
                        }
                        Command = $xml.Actions.Exec.Command
                        Arguments = $xml.Actions.Exec.Arguments
                        StartDirectory =$xml.Actions.Exec.WorkingDirectory
                        Hidden = $xml.Settings.Hidden
                    }
                }
            } Catch {
                Write-Warning ("{0}: {1}" -f $Computer, $_.Exception.Message)
            }
        }
    }
} 
Function Get-ScreenResolution {            
[void] [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")            
[void] [Reflection.Assembly]::LoadWithPartialName("System.Drawing")    
Add-Type -AssemblyName System.Windows.Forms        
$Screens = [system.windows.forms.screen]::AllScreens            

foreach ($Screen in $Screens) {            
 $DeviceName = $Screen.DeviceName            
 $Width  = $Screen.Bounds.Width            
 $Height  = $Screen.Bounds.Height            
 $IsPrimary = $Screen.Primary            

 $OutputObj = New-Object -TypeName PSobject             
 $OutputObj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $DeviceName            
 $OutputObj | Add-Member -MemberType NoteProperty -Name Width -Value $Width            
 $OutputObj | Add-Member -MemberType NoteProperty -Name Height -Value $Height            
 $OutputObj | Add-Member -MemberType NoteProperty -Name IsPrimaryMonitor -Value $IsPrimary            
 $OutputObj            

}            
}
Function Get-ScreenShot
{
    [CmdletBinding()]
    param(
        [parameter(Position  = 0, Mandatory = 0, ValueFromPipelinebyPropertyName = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$OutPath = "$env:USERPROFILE\Documents\ScreenShot",

        #screenshot_[yyyyMMdd_HHmmss_ffff].png
        [parameter(Position  = 1, Mandatory = 0, ValueFromPipelinebyPropertyName = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$FileNamePattern = 'screenshot_{0}.png',

        [parameter(Position  = 2, Mandatory = 0, ValueFromPipeline = 1, ValueFromPipelinebyPropertyName = 1)]
        [ValidateNotNullOrEmpty()]
        [int]$RepeatTimes = 0,

        [parameter(Position  = 3, Mandatory = 0, ValueFromPipelinebyPropertyName = 1)]
        [ValidateNotNullOrEmpty()]
        [int]$DurationMs = 1
     )

     begin
     {
        $ErrorActionPreference = 'Stop'
        Add-Type -AssemblyName System.Windows.Forms

        if (-not (Test-Path $OutPath))
        {
            New-Item $OutPath -ItemType Directory -Force
        }
     }

     process
     {
        1..$RepeatTimes `
        | %{
            $fileName = $FileNamePattern -f (Get-Date).ToString('yyyyMMdd_HHmmss_ffff')
            $path = Join-Path $OutPath $fileName

            $b = New-Object System.Drawing.Bitmap([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width, [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height)
            $g = [System.Drawing.Graphics]::FromImage($b)
            $g.CopyFromScreen((New-Object System.Drawing.Point(0,0)), (New-Object System.Drawing.Point(0,0)), $b.Size)
            $g.Dispose()
            $b.Save($path)

            if ($RepeatTimes -ne 0)
            {
                Start-Sleep -Milliseconds $DurationMs
            }
        }
    }
}
Function Get-ScriptAlias
{
<#
	.SYNOPSIS
		Function to retrieve the aliases inside a Powershell script file.
	
	.DESCRIPTION
		Function to retrieve the aliases inside a Powershell script file.
		Using PowerShell AST Parser we are able to retrieve the Functions and cmdlets used in the script.
	
	.PARAMETER Path
		Specifies the path of the script
	
	.EXAMPLE
		Get-ScriptAlias -Path "C:\LazyWinAdmin\testscript.ps1"
	
	.EXAMPLE
		"C:\LazyWinAdmin\testscript.ps1" | Get-ScriptAlias

    .EXAMPLE
        gci C:\LazyWinAdmin -file | Get-ScriptAlias

	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	[CmdletBinding()]
	PARAM
	(
		[Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[ValidateScript({ Test-Path -Path $_ })]
		[Alias("FullName")]
		[System.String[]]$Path
	)
	PROCESS
	{
		FOREACH ($File in $Path)
		{
			TRY
			{
				# Retrieve file content
				$ScriptContent = (Get-Content $File -Delimiter $([char]0))
				
				# AST Parsing
				$AbstractSyntaxTree = [System.Management.Automation.Language.Parser]::
				ParseInput($ScriptContent, [ref]$null, [ref]$null)
				
				# Find Aliases
				$AbstractSyntaxTree.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true) |
				ForEach-Object -Process {
					$Command = $_.CommandElements[0]
					if ($Alias = Get-Alias | Where-Object { $_.Name -eq $Command })
					{
						
						# Output information
						[PSCustomObject]@{
							File = $File
							Alias = $Alias.Name
							Definition = $Alias.Definition
							StartLineNumber = $Command.Extent.StartLineNumber
							EndLineNumber = $Command.Extent.EndLineNumber
							StartColumnNumber = $Command.Extent.StartColumnNumber
							EndColumnNumber = $Command.Extent.EndColumnNumber
							StartOffset = $Command.Extent.StartOffset
							EndOffset = $Command.Extent.EndOffset
							
						}#[PSCustomObject]
					}#if ($Alias)
				}#ForEach-Object
			}#TRY
			CATCH
			{
				Write-Error -Message $($Error[0].Exception.Message)
			} #CATCH
		}#FOREACH ($File in $Path)
	} #PROCESS
}
Function Get-ScriptDirectory
{
<#
.SYNOPSIS
   This Function retrieve the current folder path
.DESCRIPTION
   This Function retrieve the current folder path
#>
    if($hostinvocation -ne $null)
    {
        Split-Path $hostinvocation.MyCommand.path
    }
    else
    {
        Split-Path $script:MyInvocation.MyCommand.Path
    }
}
Function Get-SecurityUpdate {
    [Cmdletbinding()] 
    Param( 
        [Parameter()] 
        [String[]]$Computername = $Computername
    )              
    ForEach ($Computer in $Computername){ 
        $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
        ForEach($Path in $Paths) { 
            #Create an instance of the Registry Object and open the HKLM base key 
            Try { 
                $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer) 
            } Catch { 
                $_ 
                Continue 
            } 
            Try {
                #Drill down into the Uninstall key using the OpenSubKey Method 
                $regkey = $reg.OpenSubKey($Path)  
                #Retrieve an array of string that contain all the subkey names 
                $subkeys = $regkey.GetSubKeyNames()      
                #Open each Subkey and use GetValue Method to return the required values for each 
                ForEach ($key in $subkeys){   
                    $thisKey=$Path+"\\"+$key   
                    $thisSubKey=$reg.OpenSubKey($thisKey)   
                    # prevent Objects with empty DisplayName 
                    $DisplayName = $thisSubKey.getValue("DisplayName")
                    If ($DisplayName -AND $DisplayName -match '^Update for|rollup|^Security Update|^Service Pack|^HotFix') {
                        $Date = $thisSubKey.GetValue('InstallDate')
                        If ($Date) {
                            Write-Verbose $Date 
                            $Date = $Date -replace '(\d{4})(\d{2})(\d{2})','$1-$2-$3'
                            Write-Verbose $Date 
                            $Date = Get-Date $Date
                        } 
                        If ($DisplayName -match '(?<DisplayName>.*)\((?<KB>KB.*?)\).*') {
                            $DisplayName = $Matches.DisplayName
                            $HotFixID = $Matches.KB
                        }
                        Switch -Wildcard ($DisplayName) {
                            "Service Pack*" {$Description = 'Service Pack'}
                            "Hotfix*" {$Description = 'Hotfix'}
                            "Update*" {$Description = 'Update'}
                            "Security Update*" {$Description = 'Security Update'}
                            Default {$Description = 'Unknown'}
                        }
                        # create New Object with empty Properties 
                        $Object = [pscustomobject] @{
                            Type = $Description
                            HotFixID = $HotFixID
                            InstalledOn = $Date
                            Description = $DisplayName
                        }
                        $Object
                    } 
                }   
                $reg.Close() 
            } Catch {}                  
        }  
    }  
}
Function Get-SEPVersion { 
# All registry keys: http://www.symantec.com/business/support/index?page=content&id=HOWTO75109 
[CmdletBinding()] 
param( 
[Parameter(Position=0,Mandatory=$true,HelpMessage="Name of the computer to query SEP for", 
ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
[Alias('CN','__SERVER','IPAddress','Server')] 
[System.String] 
$ComputerName 
) 
# Create object to enable access to the months of the year 
$DateTimeFormat = New-Object System.Globalization.DateTimeFormatInfo 
#Set registry value to look for definitions path (depending on 32/64 bit OS) 
$osType=Get-WmiObject Win32_OperatingSystem -ComputerName $computername| Select OSArchitecture 
if ($osType.OSArchitecture -eq "32-bit")  
{ 
# Connect to Registry 
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName) 
# Set Registry keys to query 
$SMCKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
# Obtain Product Version value 
$SMCRegKey = $reg.opensubkey($SMCKey) 
$SEPVersion = $SMCRegKey.GetValue('ProductVersion') 
# Obtain Pattern File Date Value 
$AVRegKey = $reg.opensubkey($AVKey) 
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate') 
# Convert PatternFileDate to readable date 
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970) 
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]] 
$AVDayFileDate = [string]$AVPatternFileDate[2] 
$AVFileVersionDate = $AVDayFileDate + " " + $AVMonthFileDate + " " + $AVYearFileDate 
# Obtain Sylink Group value 
#$SylinkRegKey = $reg.opensubkey($SylinkKey) 
#$SylinkGroup = $SylinkRegKey.GetValue('CurrentGroup') 
}  
else  
{ 
# Connect to Registry 
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName) 
# Set Registry keys to query 
$SMCKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
 
# Obtain Product Version value 
$SMCRegKey = $reg.opensubkey($SMCKey) 
$SEPVersion = $SMCRegKey.GetValue('ProductVersion') 
  
# Obtain Pattern File Date Value 
$AVRegKey = $reg.opensubkey($AVKey) 
$AVPatternFileDate = $AVRegKey.GetValue("PatternFileDate") 
 
# Obtain Pattern File Date Value 
$AVRegKey = $reg.opensubkey($AVKey) 
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate') 
  
# Convert PatternFileDate to readable date 
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970) 
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]] 
$AVDayFileDate = [string]$AVPatternFileDate[2] 
$AVFileVersionDate = $AVDayFileDate + " " + $AVMonthFileDate + " " + $AVYearFileDate 
} 
$MYObject = ""| Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate 
$MYObject.ComputerName = $ComputerName 
$MYObject.SEPProductVersion = $SEPVersion 
$MYObject.SEPDefinitionDate = $AVFileVersionDate 
$MYObject 
} 
######################################################################   
# Powershell script to get the the serial numbers on remote servers   
# It will give the serial numbers on remote servers and export to csv 
# Customized script useful to every one   
# Please contact  mllsatyanarayana@gmail.com for any suggestions#   
#########################################################################  
####################serial start################# 
 
 Function Get-Serial { 
 param( 
 $computername =$env:computername 
 ) 
 
 $os = Get-WmiObject Win32_bios -ComputerName $computername -ea silentlycontinue 
 if($os){ 
 
 $SerialNumber =$os.SerialNumber 
 
 $servername=$os.PSComputerName  
  
 
  
 
 $results =new-object psobject 
 
 $results |Add-Member noteproperty SerialNumber  $SerialNumber 
 $results |Add-Member noteproperty ComputerName  $servername 
  
 
 
 #Display the results 
 
 $results | Select-Object computername,SerialNumber 
 
 } 
 
 
 else 
 
 { 
 
 $results =New-Object psobject 
 
 $results =new-object psobject 
 $results |Add-Member noteproperty SerialNumber "Na" 
 $results |Add-Member noteproperty ComputerName $servername 
 
 
  
 #display the results 
 
 $results | Select-Object computername,SerialNumber 
 
 
 
 
 } 
 
 
 
 } 
 
 $infserial =@() 
 
 
 foreach($allserver in $allservers){ 
 
$infserial += Get-Serial $allserver  
 } 
 
 $infserial  
 
 
 
 
 
<#   ####################serial end################# 
 
 
   #save the the servers in any location 
    $servers = Get-Content -Path "C:\LazyWinAdmin\Servers\servers2.txt" 
 
   foreach ($ser in $servers) 
 
   { 
    
   get-serial -computername $ser | Export-Csv -Path c:\LazyWinAdmin\Servers\serial.csv 
    
   }
   #>
    Param (
    [parameter(Mandatory=$False)]
    [ValidateSet('CORP','SVC','RES','PROD')]
    [string]$Domain = 'CORP'
)

Function Get-Servers {
    $Searcher = [adsisearcher]""
    If ($Domain = "SVC"){$searchroot = [ADSI]"LDAP://DC=svc,DC=prod,DC=vegas,DC=com"}
    elseif ($Domain = "PROD"){$searchroot = [ADSI]"LDAP://DC=prod,DC=vegas,DC=com" }
    elseif ($Domain = "RES"){$searchroot = [ADSI]"LDAP://DC=res,DC=vegas,DC=com"}
    else {$searchroot = [ADSI]"LDAP://DC=corp,DC=vegas,DC=com"}
    $Searcher.SearchRoot = $searchroot
    $Searcher.Filter = '(&(objectCategory=computer)(OperatingSystem=Windows*Server*))'
    $Searcher.pagesize = 10000
    $Searcher.sizelimit = 50000
    $Searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach-Object {$_.Properties.dnshostname}
}
# Blog Article: http://lazywinadmin.com/2014/04/powershell-getset-network-level.html
# GitHub : https://github.com/lazywinadmin/PowerShell/blob/master/TOOL-Get-Set-NetworkLevelAuthentication/Get-Set-NetworkLevelAuthentication.ps1

Function Get-NetworkLevelAuthentication
{
<#
	.SYNOPSIS
		This Function will get the NLA setting on a local machine or remote machine

	.DESCRIPTION
		This Function will get the NLA setting on a local machine or remote machine

	.PARAMETER  ComputerName
		Specify one or more computer to query

	.PARAMETER  Credential
		Specify the alternative credential to use. By default it will use the current one.
	
	.EXAMPLE
		Get-NetworkLevelAuthentication
		
		This will get the NLA setting on the localhost
	
		ComputerName     : XAVIERDESKTOP
		NLAEnabled       : True
		TerminalName     : RDP-Tcp
		TerminalProtocol : Microsoft RDP 8.0
		Transport        : tcp	

    .EXAMPLE
		Get-NetworkLevelAuthentication -ComputerName DC01
		
		This will get the NLA setting on the server DC01
	
		ComputerName     : DC01
		NLAEnabled       : True
		TerminalName     : RDP-Tcp
		TerminalProtocol : Microsoft RDP 8.0
		Transport        : tcp
	
	.EXAMPLE
		Get-NetworkLevelAuthentication -ComputerName DC01, SERVER01 -verbose
	
	.EXAMPLE
		Get-Content .\Computers.txt | Get-NetworkLevelAuthentication -verbose
		
	.NOTES
		DATE	: 2014/04/01
		AUTHOR	: Francois-Xavier Cat
		WWW		: http://lazywinadmin.com
		Twitter	: @lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the Function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				[pscustomobject][ordered]@{
					'ComputerName' = $NLAinfo.PSComputerName
					'NLAEnabled' = $NLAinfo.UserAuthenticationRequired -as [bool]
					'TerminalName' = $NLAinfo.TerminalName
					'TerminalProtocol' = $NLAinfo.TerminalProtocol
					'Transport' = $NLAinfo.transport
				}
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{
		
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}


Function Set-NetworkLevelAuthentication
{
<#
	.SYNOPSIS
		This Function will set the NLA setting on a local machine or remote machine

	.DESCRIPTION
		This Function will set the NLA setting on a local machine or remote machine

	.PARAMETER  ComputerName
		Specify one or more computers
	
	.PARAMETER EnableNLA
		Specify if the NetworkLevelAuthentication need to be set to $true or $false
	
	.PARAMETER  Credential
		Specify the alternative credential to use. By default it will use the current one.

	.EXAMPLE
		Set-NetworkLevelAuthentication -EnableNLA $true

		ReturnValue                             PSComputerName                         
		-----------                             --------------                         
		                                        XAVIERDESKTOP      
	
	.NOTES
		DATE	: 2014/04/01
		AUTHOR	: Francois-Xavier Cat
		WWW		: http://lazywinadmin.com
		Twitter	: @lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Parameter(Mandatory)]
		[Bool]$EnableNLA,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
				
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the Function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				$NLAinfo | Invoke-CimMethod -MethodName SetUserAuthenticationRequired -Arguments @{ UserAuthenticationRequired = $EnableNLA } -ErrorAction 'Continue' -ErrorVariable ErrorProcessInvokeWmiMethod
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
				if ($ErrorProcessInvokeWmiMethod) { Write-Warning -Message "PROCESS Error - $ErrorProcessInvokeWmiMethod" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{	
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}
Function Get-Skew {
<#
      .SYNOPSIS
         Gets the time of a windows server
 
      .DESCRIPTION
         Uses WMI to get the time of a remote server
 
      .PARAMETER  ServerName
         The Server to get the date and time from
 
      .EXAMPLE
         PS C:\> Get-Skew -RemoteServer RemoteServer01 -LocalServer localhost
 
      
   #>


[CmdletBinding()]
   param(
      [Parameter(Position=0, Mandatory=$true)]
      [ValidateNotNullOrEmpty()]
      [System.String]
      $Servers 
       
   )

$RemoteServer = Get-Time -ServerName $Servers
$LocalServer = Get-Time -ServerName LASDC01
 
$Skew = $LocalServer.DateTime - $RemoteServer.DateTime
 
# Check if the time is over 30 seconds
If (($Skew.TotalSeconds -gt 30) -or ($Skew.TotalSeconds -lt -30)){
   Write-Host "Time is not within 30 seconds"
} Else {
   Write-Host "Time checked ok"
}
}
Function Get-SnmpTrap {
<#
.SYNOPSIS
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
Function Get-Software {
    [OutputType('System.Software.Inventory')]
    [Cmdletbinding()] 
    Param( 
        [Parameter(ValueFromPipeline = $True,ValueFromPipelineByPropertyName = $True)] 
        [String[]]$Computername=$env:COMPUTERNAME
    )         
    Begin {
    }
    Process {     
        ForEach ($Computer in $Computername){ 
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
                ForEach($Path in $Paths) { 
                    Write-Verbose "Checking Path: $Path"
                    # Create an instance of the Registry Object and open the HKLM base key 
                    Try { 
                        $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer,'Registry64') 
                    } Catch { 
                        Write-Error $_ 
                        Continue 
                    } 
                    # Drill down into the Uninstall key using the OpenSubKey Method 
                    Try {
                        $regkey = $reg.OpenSubKey($Path)  
                        # Retrieve an array of string that contain all the subkey names 
                        $subkeys = $regkey.GetSubKeyNames()      
                        # Open each Subkey and use GetValue Method to return the required values for each 
                        ForEach ($key in $subkeys){   
                            Write-Verbose "Key: $Key"
                            $thisKey=$Path+"\\"+$key 
                            Try {  
                                $thisSubKey = $reg.OpenSubKey($thisKey)   
                                # Prevent Objects with empty DisplayName 
                                $DisplayName = $thisSubKey.getValue("DisplayName")
                                If ($DisplayName -AND $DisplayName -notmatch '^Update for|rollup|^Security Update|^Service Pack|^HotFix') {
                                    $Date = $thisSubKey.GetValue('InstallDate')
                                    If ($Date) {
                                        Try {
                                            $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)
                                        } Catch{
				                            Write-Warning "$($Computer): $_ <$($Date)>"
                                            $Date = $Null
                                        }
                                    } 
                                    # Create New Object with empty Properties 
                                    $Publisher = Try {
                                        $thisSubKey.GetValue('Publisher').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('Publisher')
                                    }
                                    $Version = Try {
                                        #Some weirdness with trailing [char]0 on some strings
                                        $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('DisplayVersion')
                                    }
                                    $UninstallString = Try {
                                        $thisSubKey.GetValue('UninstallString').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('UninstallString')
                                    }
                                    $InstallLocation = Try {
                                        $thisSubKey.GetValue('InstallLocation').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('InstallLocation')
                                    }
                                    $InstallSource = Try {
                                        $thisSubKey.GetValue('InstallSource').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('InstallSource')
                                    }
                                    $HelpLink = Try {
                                        $thisSubKey.GetValue('HelpLink').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('HelpLink')
                                    }
                                    $Object = [pscustomobject]@{
                                        Computername = $Computer
                                        DisplayName = $DisplayName
                                        Version = $Version
                                        InstallDate = $Date
                                        Publisher = $Publisher
                                        UninstallString = $UninstallString
                                        InstallLocation = $InstallLocation
                                        InstallSource = $InstallSource
                                        HelpLink = $thisSubKey.GetValue('HelpLink')
                                        EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))
                                    }
                                    $Object.pstypenames.insert(0,'System.Software.Inventory')
                                    Write-Output $Object
                                }
                            } Catch {
                                Write-Warning "$Key : $_"
                            }   
                        }
                    } Catch {}   
                    $reg.Close() 
                }                  
            } Else {
                Write-Error "$($Computer): unable to reach remote system!"
            }
        } 
    } 
} 
#Requires -Version 3.0

Function Get-StrictMode
{

    [CmdletBinding()]
    param(
        [switch]
        $showAllScopes
    )

    $currentScope = $executionContext `
    | Get-Field -name _context -valueOnly `
    | Get-Field -name _engineSessionState -valueOnly `
    | Get-Field -name currentScope -valueOnly `
    | Get-Field -name *Parent* -valueOnly `
    | Get-Field -name *Parent* -valueOnly
    
    $scope = 0
    while ($currentScope)
    {
        $strictModeVersion = $currentScope | Get-Field -name *StrictModeVersion* -valueOnly
        $currentScope = $currentScope | Get-Field -name *Parent* -valueOnly

        if ($showAllScopes)
        {
            New-Object PSObject -Property @{
                Scope             = $scope++
                StrictModeVersion = $strictModeVersion}
        }
        elseif ($strictModeVersion)
        {
            $strictModeVersion
        }
    }
}

Function Get-Field
{
    [CmdletBinding()]
    param (
        [Parameter(
            mandatory = 0,
            Position  = 0)]
        [string[]]
        $name = "*",

        [Parameter(
            mandatory = 1,
            position  = 1,
            ValueFromPipeline = 1)]
        $inputObject,
            
        [switch]
        $valueOnly
    )
 
    process
    {
        $type = $inputObject.GetType()
        [string[]]$bindingFlags = ("Public", "NonPublic", "Instance")

        $type.GetFields($bindingFlags) `
        | where {
            foreach($currentName in $name)
            {
                if ($_.Name -like $currentName)
                { 
                    return $true
                }
            }} `
        | % {
            $currentField = $_
            $currentFieldValue = $type.InvokeMember(
                $currentField.Name,
                $bindingFlags + "GetField",
                $null,
                $inputObject,
                $null
            )
                
            if ($valueOnly)
            {
                $currentFieldValue
            }
            else
            {
                $returnProperties = @{}
                foreach ($prop in @("Name", "IsPublic", "IsPrivate"))
                {
                    $ReturnProperties.$prop = $CurrentField.$prop
                }

                $returnProperties.Value = $currentFieldValue
                New-Object PSObject -Property $returnProperties
            }
        } 
    }
}


# StrictMode is Null in initilal
Get-StrictMode

# Set Strict mode to check
Set-StrictMode -Version latest

# StrictMode will show as your PS Version
Get-StrictMode
<#
Major  Minor  Build  Revision
-----  -----  -----  --------
5      0      9701   0 
#>

# turn off strict mode
Set-StrictMode -Off
<#
Major  Minor  Build  Revision
-----  -----  -----  --------
0      0      -1     -1      
#>

# StrictMode will show as 0
Get-StrictMode
<#
Major  Minor  Build  Revision
-----  -----  -----  --------
0      0      -1     -1      
#>
Function Get-StringCharCount
{
	<#
	.SYNOPSIS
		This Function will count the number of characters in a string
	.DESCRIPTION
		This Function will count the number of characters in a string
	.EXAMPLE
		PS C:\> Get-StringCharCount -String "Hello World"
	
		11
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		www.lazywinadmin.com	
	#>
	PARAM ([String]$String)
	($String -as [Char[]]).count
}
Function Get-StringLastDigit
{
<#
    .SYNOPSIS
        Get the last digit of a string
    .DESCRIPTION
        Get the last digit of a string using Regular Expression
    .PARAMETER String
        Specifies the String to check
    .EXAMPLE
        PS C:\> Get-StringLastDigit -String "Francois-Xavier.cat5"

        5
    .EXAMPLE
        PS C:\> Get-StringLastDigit -String "Francois-Xavier.cat"

        <no output>
    .EXAMPLE
        PS C:\> Get-StringLastDigit -String "Francois-Xavier.cat" -Verbose

        <no output>
        VERBOSE: The following string does not finish by a digit: Francois-Xavier.cat
    .NOTES
        Francois-Xavier Cat
        @lazywinadm
        www.lazywinadmin.com
#>
[CmdletBinding()]
PARAM($String)
    #Check if finish by Digit
    if ($String -match "^.*\d$")
    {
        # Output the last digit
        $String.Substring(($String.ToCharArray().count)-1)
    }
    else {Write-Verbose -Message "The following string does not finish by a digit: $String"}
}
Function Get-Time {
   <#
      .SYNOPSIS
         Gets the time of a windows server
 
      .DESCRIPTION
         Uses WMI to get the time of a remote server
 
      .PARAMETER  ServerName
         The Server to get the date and time from
 
      .EXAMPLE
         PS C:\> Get-Time localhost
 
      .EXAMPLE
         PS C:\> Get-Time server01.domain.local -Credential (Get-Credential)
 
   #>
   [CmdletBinding()]
   param(
      [Parameter(Position=0, Mandatory=$true)]
      [ValidateNotNullOrEmpty()]
      [System.String]
      $ServerName,
 
      $Credential
 
   )
   try {
         If ($Credential) {
            $DT = Get-WmiObject -Class Win32_LocalTime -ComputerName $servername -Credential $Credential
         } Else {
            $DT = Get-WmiObject -Class Win32_LocalTime -ComputerName $servername
         }
   }
   catch {
      throw
   }
 
   $Times = New-Object PSObject -Property @{
      ServerName = $DT.__Server
      DateTime = (Get-Date -Day $DT.Day -Month $DT.Month -Year $DT.Year -Minute $DT.Minute -Hour $DT.Hour -Second $DT.Second)
   }
   $Times
 
}
Function Get-UAC
{
<#
.Synopsis
   Check UAC configuration from Registry
.DESCRIPTION
   This cmdlet will return UAC is 'Enabled' or 'Disabled' or 'Unknown'
.EXAMPLE
   Get-UAC
.EXAMPLE
   Get-UAC -Verbose
#>

    [CmdletBinding()]
    Param
    (
    )

    begin
    {
        $path = "registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        $name = "EnableLUA"
    }

    process
    {
        $uac = Get-ItemProperty $path
        if ($uac.$name -eq 1)
        { 
            Write-Verbose ("Registry path '{0}', name '{1}', value '{2}'" -f (Get-Item $path).name, $name, $uac.$name)
            "Enabled" 
        } 
        elseif ($uac.$name -eq 0) 
        { 
            Write-Verbose ("Registry path '{0}', name '{1}', value '{2}'" -f (Get-Item $path).name, $name, $uac.$name)
            "Disabled" 
        }
        else 
        {
            Write-Verbose ("Registry path '{0}', name '{1}', value '{2}'" -f (Get-Item $path).name, $name, $uac.$name)
            "Unknown"
        }
    }
}
Function Get-UDVariable {
  get-variable | where-object {(@(
    "FormatEnumerationLimit",
    "MaximumAliasCount",
    "MaximumDriveCount",
    "MaximumErrorCount",
    "MaximumFunctionCount",
    "MaximumVariableCount",
    "PGHome",
    "PGSE",
    "PGUICulture",
    "PGVersionTable",
    "PROFILE",
    "PSSessionOption",
    "moduleBrowser",
    "psISE",
    "psUnsupportedConsoleApplications",
    "MyModules",
    "MyPowerShell",
    "MySnippets"
    

    ) -notcontains $_.name) -and `
    (([psobject].Assembly.GetType('System.Management.Automation.SpecialVariables').GetFields('NonPublic,Static') | Where-Object FieldType -eq ([string]) | ForEach-Object GetValue $null)) -notcontains $_.name
    }
}
Function Get-Uptime {
   $os = Get-WmiObject win32_operatingsystem
   $uptime = (Get-Date) - ($os.ConvertToDateTime($os.lastbootuptime))
   $Display = "Uptime: " + $Uptime.Days + " days, " + $Uptime.Hours + " hours, " + $Uptime.Minutes + " minutes" 
   Write-Output $Display
}
# End Get-Update
<#
.SYNOPSIS
Gets a URL's redirection target(s).

.DESCRIPTION
Given a URL, determines its redirection target(s), as indicated by responses
with 3xx HTTP status codes.

If the URL is not redirected, it is output as-is.

By default, the ultimate target URL is determined (if there's a chain of
redirections), but the number of redirections that are followed is limited
to 50 by default, which you may change with -MaxRedirections.

-Enumerate enumerates the redirection chain and returns an array of URLs.

.PARAMETER Url
The URL whose redirection target to determine.
You may supply multiple URLs via the pipeline.

.PARAMETER MaxRedirections
Limits the number of redirections that are followed, 50 by default.
If the limit is exceeded, a non-terminating error is reported.

.PARAMETER Enumerate
Enumerates the chain of redirections, if applicable, starting with
the input URL itself, and outputs it as an array.

If the number of actual redirections doesn't exceed the specified or default
-MaxRedirections value, the entire chain up to the ultimate target URL is
enumerated.
Otherwise, a warning is issued to indicate that the ultimate target URL wasn't
reached.

All URLs are output in absolute form, even if the targets are defined as
relative URLs.

Note that, in order to support multiple input URLs via the pipeline, each
array representing a redirection chain is output as a *single* object, so
with multiple input URLs you'll get an array of arrays as output.

.EXAMPLE
> Get-UrlRedirection http://cnn.com
http://www.cnn.com

.EXAMPLE
> Get-UrlRedirection -Enumerate http://microsoft.com/about
http://microsoft.com/about
https://microsoft.com/about
https://www.microsoft.com/about
https://www.microsoft.com/about/
https://www.microsoft.com/about/default.aspx
https://www.microsoft.com/en-us/about/

.NOTES
This Function uses the [System.Net.HttpWebRequest] .NET class and was 
inspired by http://www.powershellmagazine.com/2013/01/29/pstip-retrieve-a-redirected-url/
#>

Function Get-UrlRedirection {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory, ValueFromPipeline)] [Uri] $Url,
    [switch] $Enumerate,
    [int] $MaxRedirections = 50 # Use same default as [System.Net.HttpWebRequest]
  )

  process {
    try {

      if ($Enumerate) { # Enumerate the whole redirection chain, from input URL to ultimate target,
                        # assuming the max. count of redirects is not exceeded.
        # We must walk the chain of redirections one by one.
        # If we disallow redirections, .GetResponse() fails and we must examine
        # the exception's .Response object to get the redirect target.
        $nextUrl = $Url
        $urls = @( $nextUrl.AbsoluteUri ) # Start with the input Uri
        $ultimateFound = $false
        # Note: We add an extra loop iteration so we can determine whether
        #       the ultimate target URL was reached or not.
        foreach($i in 1..$($MaxRedirections+1)) {
          Write-Verbose "Examining: $nextUrl"
          $request = [System.Net.HttpWebRequest]::Create($nextUrl)
          $request.AllowAutoRedirect = $False
          try {
            $response = $request.GetResponse()
            # Note: In .NET *Core* the .GetResponse() for a redirected resource
            #       with .AllowAutoRedirect -eq $False throws an *exception*.
            #       We only get here on *Windows*, with the full .NET Framework.
            #       We either have the ultimate target URL, or a redirection
            #       whose target URL is reflected in .Headers['Location']
            #       !! Syntax `.Headers.Location` does NOT work.
            $nextUrlStr = $response.Headers['Location']
            $response.Close()
            # If the ultimate target URL was reached (it was already
            # recorded in the previous iteration), and if so, simply exit the loop.
            if (-not $nextUrlStr) {
              $ultimateFound = $true
              break
            }
          } catch [System.Net.WebException] {
            # The presence of a 'Location' header implies that the
            # exception must have been triggered by a HTTP redirection 
            # status code (3xx). 
            # $_.Exception.Response.StatusCode contains the specific code
            # (as an enumeration value that can be case to [int]), if needed.
            # !! Syntax `.Headers.Location` does NOT work.
            $nextUrlStr = try { $_.Exception.Response.Headers['Location'] } catch {}
            # Not being able to get a target URL implies that an unexpected
            # error ocurred: re-throw it.
            if (-not $nextUrlStr) { Throw }
          }
          Write-Verbose "Raw target: $nextUrlStr"
          if ($nextUrlStr -match '^https?:') { # absolute URL
            $nextUrl = $prevUrl = [Uri] $nextUrlStr
          } else { # URL without scheme and server component
            $nextUrl = $prevUrl = [Uri] ($prevUrl.Scheme + '://' + $prevUrl.Authority + $nextUrlStr)
          }
          if ($i -le $MaxRedirections) { $urls += $nextUrl.AbsoluteUri }          
        }
        # Output the array of URLs (chain of redirections) as a *single* object.
        Write-Output -NoEnumerate $urls
        if (-not $ultimateFound) { Write-Warning "Enumeration of $Url redirections ended before reaching the ultimate target." }

      } else { # Resolve just to the ultimate target,
                # assuming the max. count of redirects is not exceeded.

                # Note that .AllowAutoRedirect defaults to $True.
        # This will fail, if there are more redirections than the specified 
        # or default maximum.
        $request = [System.Net.HttpWebRequest]::Create($Url)
        if ($PSBoundParameters.ContainsKey('MaxRedirections')) {
          $request.MaximumAutomaticRedirections = $MaxRedirections
        }
        $response = $request.GetResponse()
        # Output the ultimate target URL.
        # If no redirection was involved, this is the same as the input URL.
        $response.ResponseUri.AbsoluteUri
        $response.Close()

       }

      } catch {
        Write-Error $_ # Report the exception as a non-terminating error.
    }
  } # process

}
Write-Host "[Math]::Round(7.9)"

Write-Host "[Convert]::ToString(576255753217, 8)"

Write-Host "[Guid]::NewGuid()"

Write-Host "[Net.Dns]::GetHostByName('schulung12')"

Write-Host "[IO.Path]::GetExtension('c:\test.txt')"

Write-Host "[IO.Path]::ChangeExtension('c:\test.txt', 'bak')"
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
 
$ErrorActionPreference="SilentlyContinue"
 
$result=@()
 
If ($Computer) {
 
Invoke-Command -ComputerName $Computer -ScriptBlock {quser} | Select-Object -Skip 1 | Foreach-Object {
 
$b=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($b[2] -like 'Disc*') {
 
$array= ([ordered]@{
'User' = $b[0]
'Computer' = $Computer
'Date' = $b[4]
'Time' = $b[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
else {
 
$array= ([ordered]@{
'User' = $b[0]
'Computer' = $Computer
'Date' = $b[5]
'Time' = $b[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
}
}
 
If ($OU) {
 
$comp=Get-ADComputer -Filter * -SearchBase "$OU" -Properties operatingsystem
 
$count=$comp.count
 
If ($count -gt 20) {
 
Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer"
 
}
 
foreach ($u in $comp) {
 
Invoke-Command -ComputerName $u.Name -ScriptBlock {quser} | Select-Object -Skip 1 | ForEach-Object {
 
$a=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($a[2] -like '*Disc*') {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[4]
'Time' = $a[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
}
 
else {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[5]
'Time' = $a[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
}
 
}
 
}
 
}
 
If ($All) {
 
$comp=Get-ADComputer -Filter * -Properties operatingsystem
 
$count=$comp.count
 
If ($count -gt 20) {
 
Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer ..."
 
}
 
foreach ($u in $comp) {
 
Invoke-Command -ComputerName $u.Name -ScriptBlock {quser} | Select-Object -Skip 1 | ForEach-Object {
 
$a=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($a[2] -like '*Disc*') {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[4]
'Time' = $a[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
else {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[5]
'Time' = $a[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
}
 
}
}
Write-Output $result
}
Function Get-UserShareDACL {
    [cmdletbinding()]
    Param(
        [Parameter()]
        $Computername = $Computername                     
    )                   
    Try {    
        Write-Verbose "Computer: $($Computername)"
        #Retrieve share information from comptuer
        $Shares = Get-WmiObject -Class Win32_LogicalShareSecuritySetting -ComputerName $Computername -ea stop
        ForEach ($Share in $Shares) {
            $MoreShare = $Share.GetRelated('Win32_Share')
            Write-Verbose "Share: $($Share.name)"
            #Try to get the security descriptor
            $SecurityDescriptor = $Share.GetSecurityDescriptor()
            #Iterate through each descriptor
            ForEach ($DACL in $SecurityDescriptor.Descriptor.DACL) {
                [pscustomobject] @{
                    Computername = $Computername
                    Name = $Share.Name
                    Path = $MoreShare.Path
                    Type = $ShareType[[int]$MoreShare.Type]
                    Description = $MoreShare.Description
                    DACLName = $DACL.Trustee.Name
                    AccessRight = $AccessMask[[int]$DACL.AccessMask]
                    AccessType = $AceType[[int]$DACL.AceType]                    
                }
            }
        }
    }
    #Catch any errors                
    Catch {}                                                    
}
Function Get-ViSession {
<#
.SYNOPSIS
Lists vCenter Sessions.

.DESCRIPTION
Lists all connected vCenter Sessions.

.EXAMPLE
PS C:\> Get-VISession

.EXAMPLE
PS C:\> Get-VISession | Where { $_.IdleMinutes -gt 5 }
#>
$SessionMgr = Get-View $DefaultViserver.ExtensionData.Client.ServiceContent.SessionManager
$AllSessions = @()
$SessionMgr.SessionList | Foreach {
$Session = New-Object -TypeName PSObject -Property @{
Key = $_.Key
UserName = $_.UserName
FullName = $_.FullName
LoginTime = ($_.LoginTime).ToLocalTime()
LastActiveTime = ($_.LastActiveTime).ToLocalTime()

}
If ($_.Key -eq $SessionMgr.CurrentSession.Key) {
$Session | Add-Member -MemberType NoteProperty -Name Status -Value “Current Session”
} Else {
$Session | Add-Member -MemberType NoteProperty -Name Status -Value “Idle”
}
$Session | Add-Member -MemberType NoteProperty -Name IdleMinutes -Value ([Math]::Round(((Get-Date) – ($_.LastActiveTime).ToLocalTime()).TotalMinutes))
$AllSessions += $Session
}
$AllSessions
}
Function Get-VMEvcMode {
<#  
.SYNOPSIS  
    Gathers information on the EVC status of a VM
.DESCRIPTION 
    Will provide the EVC status for the specified VM
.NOTES  
    Author:  Kyle Ruddy, @kmruddy, thatcouldbeaproblem.com
.PARAMETER Name
    VM name which the Function should be ran against
.EXAMPLE
	Get-VMEvcMode -Name vmName
	Retreives the EVC status of the provided VM 
#>
[CmdletBinding()] 
	param(
	[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        $Name
  	)

    Process {
        $evVM = @()

        if ($name -is [string]) {$evVM += Get-VM -Name $Name -ErrorAction SilentlyContinue}
        elseif ($name -is [array]) {

            if ($name[0] -is [string]) {
                $name | foreach {
                    $evVM += Get-VM -Name $_ -ErrorAction SilentlyContinue
                }
            }
            elseif ($name[0] -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM = $name}

        }
        elseif ($name -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM += $name}
        
        if ($evVM -eq $null) {Write-Warning "No VMs found."}
        else {
            $output = @()
            foreach ($v in $evVM) {

                $report = "" | select Name,EVCMode
                $report.Name = $v.Name
                $report.EVCMode = $v.ExtensionData.Runtime.MinRequiredEVCModeKey
                $output += $report

            }

        return $output

        }

    }

}
# Requires -Version 3.0

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


Get-WifiSSID
Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$ChineseArray -contains $_} {$CultureCode = "zh-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}
<#
    Version 6.2 Corrected to send emails from the current day and not all html in folder.
    Version 6.0
	* Added cleaning up of variables.
	* New Look and feel for the web using JCS colors.
	
    * Added Optional Parameters Days and Computer,
        - Days: It's a integer that goes from 1 to 1865 (from 1 day to 5 years behind in time to look into the logs), the default value is 3 days (if days parameter is not given)
        - Computer: In case you want to check another's computer registry.
    * The script requires that you run it in a elevated powershell console (since you're accessing the registries: security,application and system).
    * Added Log and notifications progress
    * No more infinite HTML files with the same error:
        - Added Frequency field, this is the number of repetitions of this error during the days.
        - Added FirstTime field, this is the date when the error was recorded the firsttime, during the days of query
        - Added LastTime field, ths is the date when the error was logged last time.


    How to use this script:
    EXAMPLES

    #Find all the events (Warnings and errors) for the local computer the last 3 days  (Only application and system logs)
     .\GetEventErrorsAAS.ps1

    #Find all the events (Warnings and errors) for the local computer the last 7 days (Only application and system logs)
     .\GetEventErrorsAAS.ps1 -Days 7

    #Find all events (Warnings, Errors and Information) for local computer the last 15 days (Only application and system logs)
     .\GetEventErrorsAAS.ps1 -Days 7 -AddInformation


    #Query All logs (application,system and security) for local computer in the last 3 days, will increase the time for running the script time considerably.
     .\GetEventErrorsAAS.ps1 -Days 7 -AddSecurity


     #Query all logs (application,system and security) for a remote computer 'TheRemoteServer' in the last 4 days, with informational and security info, and send the report for local email (smtp.domain.com)
     .\GetEventErrorsAAS.ps1 -Days 4 -Addinformation -AddSecurity -SendEmail -ComputersFile .\item.txt -SendEmail -computer TheRemoteServer
     

    #Added Functionality in version 6
    Attach a computer's file in txt, each computer in a new line.
    computers.txt example:
   'dc01
    ex01
    ex02
    rmte'
    
    Save the file with the name of the computers with the name "computers.txt"
     .\GetEventErrorsAAS.ps1 -ComputersFile .\computers.txt
     
     you will get 3 files in the same running script path., if you want that to be sent everyday configure a task and add the  switch -SendEmail
     SendEmail is valid for all the cases above (local or remote).
       
      .\GetEventErrorsAAS.ps1 -ComputersFile .\computers.txt -SendEmail


#>

[CmdletBinding(DefaultParameterSetName=”Computer”)]
param(
     [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=0, ParameterSetName="Computer")][Parameter(ParameterSetName='File', Position=0)]    [ValidateRange(1,1825)][int]$Days=3,
     [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=1, ParameterSetName="Computer")][ValidateLength(1,60)][string]$computer=".",
     [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=2, ParameterSetName="Computer")][Parameter(ParameterSetName='File',Position=1)][Switch]$AddInformation,
     [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=3, ParameterSetName="Computer")][Parameter(ParameterSetName='File',position=2)][Switch]$AddSecurity,
     [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=4, ParameterSetName="Computer")][Parameter(ParameterSetName='File',position=3)][switch]$SendEmail,
     [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=5, ParameterSetName="Computer")][Parameter(ParameterSetName='File',Position=4)][string]$SMTPServer="mail.domain.com",
     [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=5, ParameterSetName="File" )][string]$ComputersFile
)



#Clean Up VariableS
$CleanUpVar=@()
$CleanUpGlobal=@()

#Get start time
$TimeStart=Get-Date
$CleanUpVar+="TimeStart"

#Mail loval variables:
$mailto="joseo@lifford.com" #person or persons the would received
$mailfrom = "Reports@lifford.com" #Received from
$CleanUpVar+="mailto"
$CleanUpVar+="mailfrom"

#GLOBALs 
$global:ScriptLocation = $(get-location).Path
$global:DefaultLog = "$global:ScriptLocation\GetEventErrorsAAS.log"
$CleanUpGlobal+="ScriptLocation"
$CleanUpGlobal+="DefaultLog"

######################################################
###############       FunctionS
               ####################################JCS
#ScriptLogFunction
Function Write-Log{
    [CmdletBinding()]
    #[Alias('wl')]
    [OutputType([int])]
    Param(
            # The string to be written to the log.
            [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)] [ValidateNotNullOrEmpty()] [Alias("LogContent")] [string]$Message,
            # The path to the log file.
            [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true,Position=1)] [Alias('LogPath')] [string]$Path=$global:DefaultLog,
            [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true,Position=2)] [ValidateSet("Error","Warn","Info","Load","Execute")] [string]$Level="Info",
            [Parameter(Mandatory=$false)] [switch]$NoClobber
    )

     Process{
        
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Warning "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
            }

        # If attempting to write to a log file in a folder/path that doesn't exist
        # to create the file include path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            $NewLogFile = New-Item $Path -Force -ItemType File
            }

        else {
            # Nothing to see here yet.
            }

        # Now do the logging and additional output based on $Level
        switch ($Level) {
            'Error' {
                Write-Warning $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ERROR: `t $Message" | Out-File -FilePath $Path -Append
                }
            'Warn' {
                Write-Warning $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") WARNING: `t $Message" | Out-File -FilePath $Path -Append
                }
            'Info' {
                Write-Host $Message -ForegroundColor Green
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") INFO: `t $Message" | Out-File -FilePath $Path -Append
                }
            'Load' {
                Write-Host $Message -ForegroundColor Magenta
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") LOAD: `t $Message" | Out-File -FilePath $Path -Append
                }
            'Execute' {
                Write-Host $Message -ForegroundColor Green
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") EXEC: `t $Message" | Out-File -FilePath $Path -Append
                }
            }
    }
}
Function CheckExists{
	param(
		[Parameter(mandatory=$true,position=0)]$itemtocheck,
		[Parameter(mandatory=$true,position=1)]$colection
	)
	BEGIN{
		$item=$null
		$exist=$false
	}
	PROCESS{
		foreach($item in $colection){
			if($item.EventID -eq $itemtocheck){
				$exist=$true
				break;
			}
		}

	}
	END{
		return $exist
	}

}
Function CheckCount{
	param(
		[Parameter(mandatory=$true,position=0)]$itemtocheck,
		[Parameter(mandatory=$true,position=1)]$colection
	)
	BEGIN{
		$item=$null
		$count=0
	}
	PROCESS{
		foreach($item in $colection){
			
			if($item.EventID -eq $itemtocheck){
				$count++
			}
		}

	}
	END{
		return $count
	}

}
Function Get-Times{
	param(
		[Parameter(mandatory=$true,position=0)]$colection,
		[Parameter(mandatory=$true,position=1)]$EventID

	)
	BEGIN{
		$filterCollection= $colection | Where-Object{ $_.EventID -eq $EventID}
	}
	PROCESS{
		$previous = $filterCollection[0].TimeWritten
		$last = $filterCollection[0].TimeWritten
		foreach($item in $filterCollection){
			if($item.TimeWritten -lt $previous){
				$previous =$item.TimeWritten
			}
			if($item.TimeWritten -gt $last){
				$last = $item.TimeWritten
			}

		}

	}
	END{
		$output = New-Object psobject -Property @{
			first= $previous
			last= $last
		}
		return $output
	}

}
Function Get-EventSubscriber{
         [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,  ValueFromPipeline=$True,position=0)] [int]$Days,
        [Parameter(Mandatory=$true,  ValueFromPipeline=$True,position=1)] [ValidateSet("System","Security","Application")][ValidateNotNullOrEmpty()] [String]$LogName,
        [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=2)] [String]$computer = ".",  #dot for localhost it can be changed to get any computer in the domain (server or client)
        [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=3)] [Switch]$IncludeInfo
    )
    BEGIN{
        if($LogName -ne "Security"){
            Write-Log -Level Execute -Message "Getting $LogName Events"
        }
        else{
            Write-Log -Level Execute -Message "Getting $LogName Events. This can take a while"
        }
        #In case log is already there remove it.

    }
    
    PROCESS{
    if($LogName -ne "Security"){
        if($IncludeInfo){
            $Log=Get-EventLog -Computername $computer -LogName "$LogName" -EntryType "Information","Error","Warning" -After (Get-Date).Adddays(-$Days) | select *
        }
        else{
            $Log=Get-EventLog -Computername $computer -LogName "$LogName" -EntryType "Error","Warning" -After (Get-Date).Adddays(-$Days) | select *
        }
    }
    else{
            $Log=Get-EventLog -Computername $computer -LogName "$LogName" -EntryType FailureAudit,SuccessAudit -After (Get-Date).Adddays(-$Days) | select *
    }

    $Count = if($Log.Count){ $log.Count }else{ 1 }
   #
   # if($log.EventId -ne $null){
   #     $Count++;
   # }
   # else{
   #      
   # }

     Write-Log -Level Execute -Message "Attaching new properties to $LogName Events. Total Number of items in Log: $Count"       
     $return=@()
     $Log| foreach{$temp=$_.EventID; $valor = CheckCount -itemtocheck $temp -colection $Log;  $Dates = Get-Times -colection $Log -EventID $temp;  
        $_ |  Add-Member -Name "Frequency" -Value $valor -MemberType NoteProperty; 
        $_ |  Add-Member -Name "LastTime"  -Value $Dates.Last -MemberType NoteProperty;
        $_ |  Add-Member -Name "FirstTime" -Value $Dates.first -MemberType NoteProperty;
        $i++; $progress = ($i*100)/$Count;  
          if($progress -lt 100){Write-Progress -Activity "Attaching new properties to $LogName Events" -PercentComplete $progress; }
          else{Write-Progress -Activity "Attaching new properties to $LogName Events" -PercentComplete $progress -Complete }
          if(-not (CheckExists $temp $return)){$return+=$_ }}
        
    }
    END{
        return $return | Sort-Object Frequency -Descending
    }
}
Function ObjectsToHtml5{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$True,position=0)][String]$Computer,
        [Parameter(mandatory=$false,position=1)]$systemObjs,
        [Parameter(mandatory=$false,position=2)]$AppObjs,
        [Parameter(mandatory=$false,position=3)]$SecObjs
    )
    BEGIN{
        write-verbose "Setting Actual Date"
	    $fecha=get-date -UFormat "%Y%m%d"
	    $dia=get-date -UFormat "%A"
        
        $Fn= "$fecha$Filename"
        
        $HtmlFileName = "$global:ScriptLocation\$Filename.html"
        $title = "Event Logs $fecha/$computer"
    }
    PROCESS{
    $html= '<!DOCTYPE HTML>
<html lang="en-US">
<head>
	<meta charset="UTF-8">
	<title>'
    $html+=$title
    $html+="</title>
	<style type=""text/css"">{margin:0;padding:0}@import url(https://fonts.googleapis.com/css?family=Indie+Flower|Josefin+Sans|Orbitron:500|Yrsa);body{text-align:center;font-family:14px/1.4 'Indie Flower','Josefin Sans',Orbitron,sans-serif;font-family:'Indie Flower',cursive;font-family:'Josefin Sans',sans-serif;font-family:Orbitron,sans-serif}#page-wrap{margin:50px}tr:nth-of-type(odd){background:#eee}th{background:#EF5525;color:#fff;font-family:Orbitron,sans-serif}td,th{padding:6px;border:1px solid #ccc;text-align:center;font-size:large}table{width:90%;border-collapse:collapse;margin-left:auto;margin-right:auto;font-family:Yrsa,serif}</style>
</head>
<body>
	<h1>Event Logs Report for $computer on $dia - $fecha
</h1>
<h2> System Information </h2>
<table>
<tr>
<th>MachineName</th><th>Index</th><th>TimeGenerated</th><th>EntryType</th><th>Source</th><th>InstanceID</th><th>FirstTime</th><th>LastTime</th><th>EventID</th><th>Frequency</th><th>Message</th>
</tr>
"
foreach($item in $systemObjs){
    $machine =$item.MachineName
    $index=$item.Index
    $timeGenerated = $item.TimeGenerated
    $entrytipe =$item.EntryType
    $Source=$item.Source
    $instanceid = $item.InstanceID
    $Ft = $item.FirstTime
    $Lt = $item.LastTime
    $Eventid = $item.EventID
    $frequency = $item.Frequency
    $mensaje = $item.Message
   $html+="<tr> <td> $machine</td> <td>$index</td> <td>$timeGenerated</td> <td> $entrytipe </td> <td>$Source</td> <td>$instanceid </td><td>$Ft</td> <td>$Lt</td> <td>$Eventid</td> <td>$frequency </td> <td>$mensaje</td> </tr>"
}

$html+="
</table>
<h2> Application Information </h2>
<table>
<tr>
<th>MachineName</th><th>Index</th><th>TimeGenerated</th><th>EntryType</th><th>Source</th><th>InstanceID</th><th>FirstTime</th><th>LastTime</th><th>EventID</th><th>Frequency</th><th>Message</th>
</tr>"

foreach($item in $AppObjs){
   $machine =$item.MachineName
    $index=$item.Index
    $timeGenerated = $item.TimeGenerated
    $entrytipe =$item.EntryType
    $Source=$item.Source
    $instanceid = $item.InstanceID
    $Ft = $item.FirstTime
    $Lt = $item.LastTime
    $Eventid = $item.EventID
    $frequency = $item.Frequency
    $mensaje = $item.Message
    $html+="<tr> <td> $machine</td> <td>$index</td> <td>$timeGenerated</td> <td> $entrytipe </td> <td>$Source</td> <td>$instanceid </td><td>$Ft</td> <td>$Lt</td> <td>$Eventid</td> <td>$frequency </td> <td>$mensaje</td> </tr>"
}

$html+="
</table>
<h2> Security Information </h2>"

if($SecObjs.Count -gt 0){

$html+="
<table>
<tr>
<th>MachineName</th><th>Index</th><th>TimeGenerated</th><th>EntryType</th><th>Source</th><th>InstanceID</th><th>FirstTime</th><th>LastTime</th><th>EventID</th><th>Frequency</th><th>Message</th>
</tr>"

foreach($item in $SecObjs){
    $machine =$item.MachineName
    $index=$item.Index
    $timeGenerated = $item.TimeGenerated
    $entrytipe =$item.EntryType
    $Source=$item.Source
    $instanceid = $item.InstanceID
    $Ft = $item.FirstTime
    $Lt = $item.LastTime
    $Eventid = $item.EventID
    $frequency = $item.Frequency
    $mensaje = $item.Message
    $html+="<tr> <td> $machine</td> <td>$index</td> <td>$timeGenerated</td> <td> $entrytipe </td> <td>$Source</td> <td>$instanceid </td><td>$Ft</td> <td>$Lt</td> <td>$Eventid</td> <td>$frequency </td> <td>$mensaje</td> </tr>"

}
$html+="</table>"

}
else{

$html+="<p> Security objects not selected in the query. If you want this information please re-run the script with the option '-AddSecurity' <br> Also remeber that you can also add the information information with the switch '-Addinformation'</p>"
}


$html+="
	<footer>
	<a href=""https://www.j0rt3g4.com"" target=""_blank"">
	2017 - J0rt3g4 Consulting Services </a> | - &#9400; All rigths reserved.
	</footer>
</body>
</html>"

    
    }
    END{
        $html | Out-File "$global:ScriptLocation\$fecha-$dia-$computername.html" 
    
    }
}
#Get warnings and info on each event viewer log.
Function GetEventErrors{
<#
  .SYNOPSIS
    Get Warning and Errors from event viewer logs in local computer for the last 3 days (by default) it can be extended to any amount of days.
  .DESCRIPTION
  .EXAMPLE
  GetLogHTML ScriptPathVariable NDays
  .PARAMETERS 
  #>
  [CmdletBinding()]
   param
  (
    [Parameter(Mandatory=$true,  ValueFromPipeline=$True,position=0)] [String]$ScriptPath=".",
	[Parameter(Mandatory=$true,  ValueFromPipeline=$True,position=1)] [int]$Days=$global:DefaultNumberOfDays, #Day(s) behind for the checking of the logs (default 3) set in line 50
	[Parameter(Mandatory=$false, ValueFromPipeline=$True,position=2)] [String]$computer = ".",  #dot for localhost it can be changed to get any computer in the domain (server or client)
    [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=3)] [Switch]$AddInfo=$false, #Add Information Events
    [Parameter(Mandatory=$false, ValueFromPipeline=$True,position=4)] [Switch]$AddSecu=$false  #add security Log
  )

  BEGIN{
	write-Verbose "Preparing script's Variables"
	#SimpleLogname
	$SystemLogName = "System" #other options are: security, application, forwarded events
	$AppLogname= "Application"
	$SecurityLogName= "Security"    
	#set html header in a variable CSS3
	#$header= "<style type=""text/css"">body,html{height:100%}a,abbr,acronym,address,applet,b,big,blockquote,body,caption,center,cite,code,dd,del,dfn,div,dl,dt,em,fieldset,font,form,html,i,iframe,img,ins,kbd,label,legend,li,object,ol,p,pre,q,s,samp,small,span,strike,strong,sub,sup,table,tbody,td,tfoot,th,thead,tr,tt,u,ul,var{margin:0;padding:0;border:0;outline:0;font-size:100%;vertical-align:baseline;background:0 0}body{line-height:1}ol,ul{list-style:none}blockquote,q{quotes:none}blockquote:after,blockquote:before,q:after,q:before{content:'';content:none}:focus{outline:0}del{text-decoration:line-through}table{border-spacing:0;margin: 50px 0px 50px 10px;}body{font-family:Arial,Helvetica,sans-serif;margin:0 15px;width:520px}a:link,a:visited{color:#666;font-weight:700;text-decoration:none}a:active,a:hover{color:#bd5a35;text-decoration:underline}table a:link{color:#666;font-weight:700;text-decoration:none}table a:visited{color:#999;font-weight:700;text-decoration:none}table a:active,table a:hover{color:#bd5a35;text-decoration:underline}table{font-family:Arial,Helvetica,sans-serif;color:#666;font-size:12px;text-shadow:1px 1px 0 #fff;background:#eaebec;border:1px solid #ccc;-moz-border-radius:3px;-webkit-border-radius:3px;border-radius:3px;-moz-box-shadow:0 1px 2px #d1d1d1;-webkit-box-shadow:0 1px 2px #d1d1d1;box-shadow:0 1px 2px #d1d1d1}table th{padding:21px 25px 22px;border-top:1px solid #fafafa;border-bottom:1px solid #e0e0e0;background:#ededed;background:-webkit-gradient(linear,left top,left bottom,from(#ededed),to(#ebebeb));background:-moz-linear-gradient(top,#ededed,#ebebeb)}table th:first-child{text-align:left;padding-left:20px}table tr:first-child th:first-child{-moz-border-radius-topleft:3px;-webkit-border-top-left-radius:3px;border-top-left-radius:3px}table tr:first-child th:last-child{-moz-border-radius-topright:3px;-webkit-border-top-right-radius:3px;border-top-right-radius:3px}table tr{text-align:center;padding-left:20px}table tr td:first-child{text-align:left;padding-left:20px;border-left:0}table tr td{padding:18px;border-top:1px solid #fff;border-bottom:1px solid #e0e0e0;border-left:1px solid #e0e0e0;background:#fafafa;background:-webkit-gradient(linear,left top,left bottom,from(#fbfbfb),to(#fafafa));background:-moz-linear-gradient(top,#fbfbfb,#fafafa)}table tr.even td{background:#f6f6f6;background:-webkit-gradient(linear,left top,left bottom,from(#f8f8f8),to(#f6f6f6));background:-moz-linear-gradient(top,#f8f8f8,#f6f6f6)}table tr:last-child td{border-bottom:0}table tr:last-child td:first-child{-moz-border-radius-bottomleft:3px;-webkit-border-bottom-left-radius:3px;border-bottom-left-radius:3px}table tr:last-child td:last-child{-moz-border-radius-bottomright:3px;-webkit-border-bottom-right-radius:3px;border-bottom-right-radius:3px}table tr:hover td{background:#f2f2f2;background:-webkit-gradient(linear,left top,left bottom,from(#f2f2f2),to(#f0f0f0));background:-moz-linear-gradient(top,#f2f2f2,#f0f0f0);div{font-size:20px;}}</style>";
	#$header= "<style type=""text/css"">{margin:0;padding:0}body{font:14px/1.4 Georgia,Serif}#page-wrap{margin:50px}p{margin:20px 0}table{width:100%;border-collapse:collapse}tr:nth-of-type(odd){background:#eee}th{background:#333;color:#fff;font-weight:700}td,th{padding:6px;border:1px solid #ccc;text-align:left}</style>";
	$header= "<style type=""text/css"">{margin:0;padding:0}@import url(https://fonts.googleapis.com/css?family=Indie+Flower|Josefin+Sans|Orbitron:500|Yrsa);body{font-family:14px/1.4 'Indie Flower','Josefin Sans',Orbitron,sans-serif;font-family:'Indie Flower',cursive;font-family:'Josefin Sans',sans-serif;font-family:Orbitron,sans-serif}#page-wrap{margin:50px}tr:nth-of-type(odd){background:#eee}th{background:#EF5525;color:#fff;font-family:Orbitron,sans-serif}td,th{padding:6px;border:1px solid #ccc;text-align:center;font-size:large}table{width:90%;border-collapse:collapse;margin-left:auto;margin-right:auto;font-family:Yrsa,serif}</style>";
  }
  PROCESS{
	#GET ALL ITEMS in Event Viewer with the selected options
	if(-not $AddSecu -and -not $AddInfo){
        Write-Log -Level Load -Message "Querying Logs in System and Application Logs without informational items (just Warnings and errors)"
        $system= Get-EventSubscriber -Days $Days -LogName "$SystemLogName" -computer $computerName 
        $appl= Get-EventSubscriber -Days $Days -LogName "$AppLogname" -computer $computerName
    }
    elseif($AddSecu -and -not $AddInfo){
        Write-Log -Level Load -Message "Querying Logs in System, Application and security Logs without informational items (just Warnings and errors)"
        $system= Get-EventSubscriber -Days $Days -LogName "$SystemLogName" -computer $computerName 
        $appl= Get-EventSubscriber -Days $Days -LogName "$AppLogname" -computer $computerName
        $security= Get-EventSubscriber -Days $Days -LogName "$SecurityLogName" -computer $computerName
    }
    elseif(-not $AddSecu -and $AddInfo){
        Write-Log -Level Load -Message "Querying Logs in System and Application Logs WITH informational items"
        $system= Get-EventSubscriber -Days $Days -LogName "$SystemLogName" -computer $computerName -IncludeInfo
        $appl= Get-EventSubscriber -Days $Days -LogName "$AppLogname" -computer $computerName -IncludeInfo
    }
    else{
        Write-Log -Level Load -Message "Querying Logs in System, Application and security Logs WITH informational items (just in system and application, security doesn't have informational items)"
        $system= Get-EventSubscriber -Days $Days -LogName "$SystemLogName" -computer $computerName -IncludeInfo
        $appl= Get-EventSubscriber -Days $Days -LogName "$AppLogname" -computer $computerName -IncludeInfo
        $security= Get-EventSubscriber -Days $Days -LogName "$SecurityLogName" -computer $computerName
    }

     if($AddSecu){
        ObjectsToHtml5 -Computer $computerName -systemObjs $system -AppObjs $appl -SecObj $security
     }
     else{
        ObjectsToHtml5 -Computer $computerName -systemObjs $system -AppObjs $appl
     }
  }
  END{
    write-verbose "Done Exporting"
  }
 }
Function ShowTimeMS{
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline=$True,position=0,mandatory=$true)]	[datetime]$timeStart,
	[Parameter(ValueFromPipeline=$True,position=1,mandatory=$true)]	[datetime]$timeEnd
  )
  BEGIN {}
  PROCESS{
	write-Verbose "Stamping time"
	$diff=New-TimeSpan $TimeStart $TimeEnd
	Write-verbose "Timediff= $diff"
    
    if($diff.TotalMinutes -gt 60){
        #hours
        $hours = $diff.TotalHours
        Write-Log -Level Info -Message "End Script in $hours hour(s)"
    }
    elseif($diff.TotalSeconds -gt 60){
        #minutes
        $minutes = $diff.TotalMinutes
        Write-Log -Level Info -Message "End Script in $minutes minute(s)"
    }
    elseif($diff.TotalMilliseconds -gt 100){
        #seconds
        $seconds = $diff.TotalSeconds
        Write-Log -Level Info -Message "End Script in $seconds seconds"
    }
    else{
        #ms
        $miliseconds = $diff.TotalMilliseconds
        Write-Log -Level Info -Message "End Script in $miliseconds miliseconds"
    }
  }
  END{}
}
#get script directory
Write-Log -Level Info -Message "*******************************     Start Script     ******************************"


if($AddSecurity){
    Write-Log -Level Warn -Message "Using the ""AddSecurity"" switch will increase the time of execution of the script"
    $key = Read-Host "Are you sure you want to continue?(Y/N) "
    $CleanUpVar+="key"
    if($key -ne "Y" -or $key -ne "y"){
        $TimeEnd=Get-Date
		$CleanUpVar+="TimeEnd"
        #Write export total time into console
        $time=ShowTimeMS $TimeStart $TimeEnd 
        $CleanUpVar+="time"
        Write-Log -Level Info -Message "End Script in $time miliseconds"
        exit(0)
    }
}




#call the eventlog Function and export the info to html (the 3 at the end is the number of days backwards where it will search, using 1 -> last 24 hours, 2 ->48 days, etc)

if($computer -eq "."){
    $computerName = $env:computername
}
else{
	$computerName = $computer
}

$CleanUpVar+="computerName"

if([string]::IsNullOrEmpty($ComputersFile) ){
    Write-Log -Level Execute -Message "Creating Html Files"
    
    if(-not $AddInformation -and -not $AddSecurity){
        GetEventErrors -ScriptPath $global:ScriptLocation -Days $Days -computer $computerName
    }
    elseif(-not $AddInformation -and $AddSecurity){
        GetEventErrors -ScriptPath $global:ScriptLocation -Days $Days -computer $computerName  -AddInfo:$false -AddSecu
    }
    elseif( $AddInformation -and -not $AddSecurity){
        GetEventErrors -ScriptPath $global:ScriptLocation -Days $Days -computer $computerName  -AddInfo
    }
    else{
        GetEventErrors -ScriptPath $global:ScriptLocation -Days $Days -computer $computerName  -AddInfo -AddSecu
    }
    
    if($SendEmail){
        $fecha=get-date -UFormat "%Y%m%d"
        $dia=get-date -UFormat "%A"
        $Subject="Report from $computer on $fecha" 
        $HtmlFileName ="$global:ScriptLocation\$fecha-$dia-$computerName.html" 

        $CleanUpVar+="fecha"
        $CleanUpVar+="dia"
        $CleanUpVar+="subject"
        $CleanUpVar+="HtmlFileName"
        Send-MailMessage -From $mailfrom -To $mailto -Subject  $Subject -Body "JCS $Subject" -Attachments "$HtmlFileName" -Priority High -dno onSuccess, onFailure -SmtpServer $SMTPServer
    }

}
else{
    $computers = Get-Content $ComputersFile

    foreach($computer in $computers){

    	if($computer -eq "."){
		    $computerName = $env:computername
	    }
	    else{
		    $computerName = $computer
	    }
	    

        Write-Log -Level Load -Message "Looking information for computer $computerName"

        if(-not $AddInformation -and -not $AddSecurity){
            GetEventErrors -ScriptPath $global:ScriptLocation -Days $Days -computer $computerName 
        }
        elseif(-not $AddInformation -and $AddSecurity){
            GetEventErrors -ScriptPath $global:ScriptLocation -Days $Days -computer $computerName -AddInfo:$false -AddSecu
        }
        elseif( $AddInformation -and -not $AddSecurity){
            GetEventErrors -ScriptPath $global:ScriptLocation -Days $Days -computer $computerName -AddInfo
        }
        else{
            GetEventErrors -ScriptPath $global:ScriptLocation -Days $Days -computer $computerName -AddInfo -AddSecu
        }
    }
    if($SendEmail){
        $fecha=get-date -UFormat "%Y%m%d"
        $dia=get-date -UFormat "%A"
        $Subject="Report from several computers on $fecha" 
        #Get all Html files inside that folder
        $files= [System.IO.Directory]::GetFiles("$global:ScriptLocation", "*.html", [System.IO.SearchOption]::AllDirectories);
        $todayFiles = files | where{ $_ -match $fecha}
        $CleanUpVar+="fecha"
        $CleanUpVar+="dia"
        $CleanUpVar+="subject"
       
        Send-MailMessage -From $mailfrom -To $mailto -Subject  $Subject -Body "JCS $Subject" -Attachments $todayFiles -Priority High -dno onSuccess, onFailure -SmtpServer $SMTPServer
     }
}


$TimeEnd=Get-Date
$time=ShowTimeMS $TimeStart $TimeEnd 
Write-Log -Level Info -Message "******************************    Finished Script     *****************************"
#get the info for finish
$CleanUpVar| ForEach-Object{
	Remove-Variable $_
}
$CleanUpGlobal | ForEach-Object{
	Remove-Variable -Scope global $_
}
Remove-Variable CleanUpGlobal,CleanUpVar
<##############################################################################
Ashley McGlone
Microsoft Premier Field Engineer
http://aka.ms/GoateePFE
May 2015

This script includes the following Functions:
Get-GPLink
Get-GPUnlinked
Copy-GPRegistryValue

All code has been tested on Windows Server 2008 R2 with PowerShell v2.0.

Requires:
-PowerShell v2 or above
-RSAT
-ActiveDirectory module
-GroupPolicy module

See the end of this file for sample usage.
Press F5 to run the script and only put the Functions into memoory.
The BREAK statement keeps the sample code from running.
Edit and highlight the sample code.  Then run it with F8.

See the code below for comments and documentation inline.


LEGAL DISCLAIMER
This Sample Code is provided for the purpose of illustration only and is not
intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
nonexclusive, royalty-free right to use and modify the Sample Code and to
reproduce and distribute the object code form of the Sample Code, provided
that You agree: (i) to not use Our name, logo, or trademarks to market Your
software product in which the Sample Code is embedded; (ii) to include a valid
copyright notice on Your software product in which the Sample Code is embedded;
and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
against any claims or lawsuits, including attorneys’ fees, that arise or result
from the use or distribution of the Sample Code.
 
This posting is provided "AS IS" with no warranties, and confers no rights. Use
of included script samples are subject to the terms specified
at http://www.microsoft.com/info/cpyright.htm.
##############################################################################>


<#
.SYNOPSIS
This Function creates a report of all group policy links, their locations, and
their configurations in the current domain.  Output is a CSV file.
.DESCRIPTION
Long description
.PARAMETER Path
Optional parameter.  If specified, it will return GPLinks for a specific OU or domain root rather than all GPLinks.
.EXAMPLE
Get-GPLink | Out-GridView
.EXAMPLE
Get-GPLink -Path 'OU=Users,OU=IT,DC=wingtiptoys,DC=local' | Out-GridView
.EXAMPLE
Get-GPLink -Path 'DC=wingtiptoys,DC=local' | Out-GridView
.EXAMPLE
Get-GPLink -Path 'DC=wingtiptoys,DC=local' | ForEach-Object {$_.DisplayName}
.NOTES
For more information on gPLink, gPOptions, and gPLinkOptions see:
 [MS-GPOL]: Group Policy: Core Protocol
  http://msdn.microsoft.com/en-us/library/cc232478.aspx
 2.2.2 Domain SOM Search
  http://msdn.microsoft.com/en-us/library/cc232505.aspx
 2.3 Directory Service Schema Elements
  http://msdn.microsoft.com/en-us/library/cc422909.aspx
 3.2.5.1.5 GPO Search
  http://msdn.microsoft.com/en-us/library/cc232537.aspx

SOM is an acronym for Scope of Management, referring to any location where
a group policy could be linked: domain, OU, site.

This GPO report does not list GPO filtering by permissions.

Helpful commands when inspecting GPO links:
Get-ADOrganizationalUnit -Filter {Name -eq 'Production'} | Select-Object -ExpandProperty LinkedGroupPolicyObjects
Get-ADOrganizationalUnit -Filter * | Select-Object DistinguishedName, LinkedGroupPolicyObjects
Get-ADObject -Identity 'OU=HR,DC=wingtiptoys,DC=local' -Property gPLink
#>
Function Get-GPLink {
Param(
    [Parameter()]
    [string]
    $Path
)

    # Requires RSAT installed and features enabled
    Import-Module GroupPolicy
    Import-Module ActiveDirectory

    # Pick a DC to target
    $Server = Get-ADDomainController -Discover | Select-Object -ExpandProperty HostName

    # Grab a list of all GPOs
    $GPOs = Get-GPO -All -Server $Server | Select-Object ID, Path, DisplayName, GPOStatus, WMIFilter, CreationTime, ModificationTime, User, Computer

    # Create a hash table for fast GPO lookups later in the report.
    # Hash table key is the policy path which will match the gPLink attribute later.
    # Hash table value is the GPO object with properties for reporting.
    $GPOsHash = @{}
    ForEach ($GPO in $GPOs) {
        $GPOsHash.Add($GPO.Path,$GPO)
    }

    # Empty array to hold all possible GPO link SOMs
    $gPLinks = @()

    If ($PSBoundParameters.ContainsKey('Path')) {

        $gPLinks += `
         Get-ADObject -Server $Server -Identity $Path -Properties name, distinguishedName, gPLink, gPOptions |
         Select-Object name, distinguishedName, gPLink, gPOptions

    } Else {

        # GPOs linked to the root of the domain
        #  !!! Get-ADDomain does not return the gPLink attribute
        $gPLinks += `
         Get-ADObject -Server $Server -Identity (Get-ADDomain).distinguishedName -Properties name, distinguishedName, gPLink, gPOptions |
         Select-Object name, distinguishedName, gPLink, gPOptions

        # GPOs linked to OUs
        #  !!! Get-GPO does not return the gPLink attribute
        $gPLinks += `
         Get-ADOrganizationalUnit -Server $Server -Filter * -Properties name, distinguishedName, gPLink, gPOptions |
         Select-Object name, distinguishedName, gPLink, gPOptions

        # GPOs linked to sites
        $gPLinks += `
         Get-ADObject -Server $Server -LDAPFilter '(objectClass=site)' -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" -SearchScope OneLevel -Properties name, distinguishedName, gPLink, gPOptions |
         Select-Object name, distinguishedName, gPLink, gPOptions
    }

    # Empty report array
    $report = @()

    # Loop through all possible GPO link SOMs collected
    ForEach ($SOM in $gPLinks) {
        # Filter out policy SOMs that have a policy linked
        If ($SOM.gPLink) {

            # If an OU has 'Block Inheritance' set (gPOptions=1) and no GPOs linked,
            # then the gPLink attribute is no longer null but a single space.
            # There will be no gPLinks to parse, but we need to list it with BlockInheritance.
            If ($SOM.gPLink.length -gt 1) {
                # Use @() for force an array in case only one object is returned (limitation in PS v2)
                # Example gPLink value:
                #   [LDAP://cn={7BE35F55-E3DF-4D1C-8C3A-38F81F451D86},cn=policies,cn=system,DC=wingtiptoys,DC=local;2][LDAP://cn={046584E4-F1CD-457E-8366-F48B7492FBA2},cn=policies,cn=system,DC=wingtiptoys,DC=local;0][LDAP://cn={12845926-AE1B-49C4-A33A-756FF72DCC6B},cn=policies,cn=system,DC=wingtiptoys,DC=local;1]
                # Split out the links enclosed in square brackets, then filter out
                # the null result between the closing and opening brackets ][
                $links = @($SOM.gPLink -split {$_ -eq '[' -or $_ -eq ']'} | Where-Object {$_})
                # Use a for loop with a counter so that we can calculate the precedence value
                For ( $i = $links.count - 1 ; $i -ge 0 ; $i-- ) {
                    # Example gPLink individual value (note the end of the string):
                    #   LDAP://cn={7BE35F55-E3DF-4D1C-8C3A-38F81F451D86},cn=policies,cn=system,DC=wingtiptoys,DC=local;2
                    # Splitting on '/' and ';' gives us an array every time like this:
                    #   0: LDAP:
                    #   1: (null value between the two //)
                    #   2: distinguishedName of policy
                    #   3: numeric value representing gPLinkOptions (LinkEnabled and Enforced)
                    $GPOData = $links[$i] -split {$_ -eq '/' -or $_ -eq ';'}
                    # Add a new report row for each GPO link
                    $report += New-Object -TypeName PSCustomObject -Property @{
                        Name              = $SOM.Name;
                        OUDN              = $SOM.distinguishedName;
                        PolicyDN          = $GPOData[2];
                        Precedence        = $links.count - $i
                        GUID              = "{$($GPOsHash[$($GPOData[2])].ID)}";
                        DisplayName       = $GPOsHash[$GPOData[2]].DisplayName;
                        GPOStatus         = $GPOsHash[$GPOData[2]].GPOStatus;
                        WMIFilter         = $GPOsHash[$GPOData[2]].WMIFilter.Name;
                        GPOCreated        = $GPOsHash[$GPOData[2]].CreationTime;
                        GPOModified       = $GPOsHash[$GPOData[2]].ModificationTime;
                        UserVersionDS     = $GPOsHash[$GPOData[2]].User.DSVersion;
                        UserVersionSysvol = $GPOsHash[$GPOData[2]].User.SysvolVersion;
                        ComputerVersionDS = $GPOsHash[$GPOData[2]].Computer.DSVersion;
                        ComputerVersionSysvol = $GPOsHash[$GPOData[2]].Computer.SysvolVersion;
                        Config            = $GPOData[3];
                        LinkEnabled       = [bool](!([int]$GPOData[3] -band 1));
                        Enforced          = [bool]([int]$GPOData[3] -band 2);
                        BlockInheritance  = [bool]($SOM.gPOptions -band 1)
                    } # End Property hash table
                } # End For
            }
        }
    } # End ForEach

    # Output the results to CSV file for viewing in Excel
    $report |
     Select-Object OUDN, BlockInheritance, LinkEnabled, Enforced, Precedence, `
      DisplayName, GPOStatus, WMIFilter, GUID, GPOCreated, GPOModified, `
      UserVersionDS, UserVersionSysvol, ComputerVersionDS, ComputerVersionSysvol, PolicyDN
}

<#########################################################################sdg#>

<#
.SYNOPSIS
Used to discover GPOs that are not linked anywhere in the domain.
.DESCRIPTION
All GPOs in the domain are returned. The Linked property indicates true if any links exist.  The property is blank if no links exist.
.EXAMPLE
Get-GPUnlinked | Out-GridView
.EXAMPLE
Get-GPUnlinked | Where-Object {!$_.Linked} | Out-GridView
.NOTES
This Function does not look for GPOs linked to sites.
Use the Get-GPLink Function to view those.
#>
Function Get-GPUnlinked {

    Import-Module GroupPolicy
    Import-Module ActiveDirectory

    # BUILD LIST OF ALL POLICIES IN A HASH TABLE FOR QUICK LOOKUP
    $AllPolicies = Get-ADObject -Filter * -SearchBase "CN=Policies,CN=System,$((Get-ADDomain).Distinguishedname)" -SearchScope OneLevel -Property DisplayName, whenCreated, whenChanged
    $GPHash = @{}
    ForEach ($Policy in $AllPolicies) {
        $GPHash.Add($Policy.DistinguishedName,$Policy)
    }

    # BUILD LIST OF ALL LINKED POLICIES
    $AllLinkedPolicies = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty LinkedGroupPolicyObjects -Unique
    $AllLinkedPolicies += Get-ADDomain | Select-Object -ExpandProperty LinkedGroupPolicyObjects -Unique

    # FLAG EACH ONE WITH A LINKED PROPERTY
    ForEach ($Policy in $AllLinkedPolicies) {
        $GPHash[$Policy].Linked = $true
    }

    # POLICY LINKED STATUS
    $GPHash.Values | Select-Object whenCreated, whenChanged, Linked, DisplayName, Name, DistinguishedName

    ### NOTE THAT whenChanged IS NOT A REPLICATED VALUE
}

<#########################################################################sdg#>


# HELPER Function FOR Copy-GPRegistryValue
Function DownTheRabbitHole {
[CmdletBinding()]
Param(
    [Parameter()]
    [String[]]
    $rootPaths,
    [Parameter()]
    [String]
    $SourceGPO,
    [Parameter()]
    [String]
    $DestinationGPO
)

    $ErrorActionPreference = 'Continue'

    ForEach ($rootPath in $rootPaths) {

        Write-Verbose "SEARCHING PATH [$SourceGPO] [$rootPath]"
        Try {
            $children = Get-GPRegistryValue -Name $SourceGPO -Key $rootPath -Verbose -ErrorAction Stop
        }
        Catch {
            Write-Warning "REGISTRY PATH NOT FOUND [$SourceGPO] [$rootPath]"
            $children = $null
        }

        $Values = $children | Where-Object {-not [string]::IsNullOrEmpty($_.PolicyState)}
        If ($Values) {
            ForEach ($Value in $Values) {
                If ($Value.PolicyState -eq "Delete") {
                    Write-Verbose "SETTING DELETE [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)]"
                    If ([string]::IsNullOrEmpty($_.Valuename)) {
                        Write-Warning "EMPTY VALUENAME, POTENTIAL SETTING FAILURE, CHECK MANUALLY [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)]"
                        Set-GPRegistryValue -Disable -Name $DestinationGPO -Key $Value.FullKeyPath -Verbose | Out-Null
                    } Else {

                        # Warn if overwriting an existing value in the DestinationGPO.
                        # This usually does not get triggered for DELETE settings.
                        Try {
                            $OverWrite = $true
                            $AlreadyThere = Get-GPRegistryValue -Name $DestinationGPO -Key $rootPath -ValueName $Value.Valuename -Verbose -ErrorAction Stop
                        }
                        Catch {
                            $OverWrite = $false
                        }
                        Finally {
                            If ($OverWrite) {
                                Write-Warning "OVERWRITING PREVIOUS VALUE [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)] [$($AlreadyThere.Value -join ';')]"
                            }
                        }

                        Set-GPRegistryValue -Disable -Name $DestinationGPO -Key $Value.FullKeyPath -ValueName $Value.Valuename -Verbose | Out-Null
                    }
                } Else {
                    # PolicyState = "Set"
                    Write-Verbose "SETTING SET [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)]"

                    # Warn if overwriting an existing value in the DestinationGPO.
                    # This can occur when consolidating multiple GPOs that may define the same setting, or when re-running a copy.
                    # We do not check to see if the values match.
                    Try {
                        $OverWrite = $true
                        $AlreadyThere = Get-GPRegistryValue -Name $DestinationGPO -Key $rootPath -ValueName $Value.Valuename -Verbose -ErrorAction Stop
                    }
                    Catch {
                        $OverWrite = $false
                    }
                    Finally {
                        If ($OverWrite) {
                            Write-Warning "OVERWRITING PREVIOUS VALUE [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)] [$($AlreadyThere.Value -join ';')]"
                        }
                    }

                    $Value | Set-GPRegistryValue -Name $DestinationGPO -Verbose | Out-Null
                }
            }
        }
                
        $subKeys = $children | Where-Object {[string]::IsNullOrEmpty($_.PolicyState)} | Select-Object -ExpandProperty FullKeyPath
        if ($subKeys) {
            DownTheRabbitHole -rootPaths $subKeys -SourceGPO $SourceGPOSingle -DestinationGPO $DestinationGPO -Verbose
        }
    }
}


<#
.SYNOPSIS
Copies GPO registry settings from one or more policies to another.
.DESCRIPTION
Long description
.PARAMETER Mode
Indicates which half of the GPO settings to copy.  Three possible values: All, User, Computer.
.PARAMETER SourceGPO
Display name of one or more GPOs from which to copy settings.
.PARAMETER DestinationGPO
Display name of destination GPO to receive the settings.
If the destination GPO does not exist, then it creates it.
.EXAMPLE
Copy-GPRegistryValue -Mode All -SourceGPO "IE Test" -DestinationGPO "NewMergedGPO" -Verbose
.EXAMPLE
Copy-GPRegistryValue -Mode All -SourceGPO "foo", "Starter User", "Starter Computer" -DestinationGPO "NewMergedGPO" -Verbose
.EXAMPLE
Copy-GPRegistryValue -Mode User -SourceGPO 'User Settings' -DestinationGPO 'New Merged GPO' -Verbose
.EXAMPLE
Copy-GPRegistryValue -Mode Computer -SourceGPO 'Computer Settings' -DestinationGPO 'New Merged GPO' -Verbose
.NOTES
Helpful commands when inspecting GPO links:
Get-ADOrganizationalUnit -Filter {Name -eq 'Production'} | Select-Object -ExpandProperty LinkedGroupPolicyObjects
Get-ADOrganizationalUnit -Filter * | Select-Object DistinguishedName, LinkedGroupPolicyObjects
Get-ADObject -Identity 'OU=HR,DC=wingtiptoys,DC=local' -Property gPLink
#>
Function Copy-GPRegistryValue {
[CmdletBinding()]
Param(
    [Parameter()]
    [ValidateSet('All','User','Computer')]
    [String]
    $Mode = 'All',
    [Parameter()]
    [String[]]
    $SourceGPO,
    [Parameter()]
    [String]
    $DestinationGPO
)
    Import-Module GroupPolicy -Verbose:$false

    $ErrorActionPreference = 'Continue'

    Switch ($Mode) {
        'All'      {$rootPaths = "HKCU\Software","HKLM\System","HKLM\Software"; break}
        'User'     {$rootPaths = "HKCU\Software"                              ; break}
        'Computer' {$rootPaths = "HKLM\System","HKLM\Software"                ; break}
    }
    
    If (Get-GPO -Name $DestinationGPO -ErrorAction SilentlyContinue) {
        Write-Verbose "DESTINATION GPO EXISTS [$DestinationGPO]"
    } Else {
        Write-Verbose "CREATING DESTINATION GPO [$DestinationGPO]"
        New-GPO -Name $DestinationGPO -Verbose | Out-Null
    }

    $ProgressCounter = 0
    $ProgressTotal   = @($SourceGPO).Count   # Syntax for PSv2 compatibility
    ForEach ($SourceGPOSingle in $SourceGPO) {

        Write-Progress -PercentComplete ($ProgressCounter / $ProgressTotal * 100) -Activity "Copying GPO settings to: $DestinationGPO" -Status "From: $SourceGPOSingle"

        If (Get-GPO -Name $SourceGPOSingle -ErrorAction SilentlyContinue) {

            Write-Verbose "SOURCE GPO EXISTS [$SourceGPOSingle]"

            DownTheRabbitHole -rootPaths $rootPaths -SourceGPO $SourceGPOSingle -DestinationGPO $DestinationGPO -Verbose

            Get-GPOReport -Name $SourceGPOSingle -ReportType Xml -Path "$pwd\report_$($SourceGPOSingle).xml"
            $nonRegistry = Select-String -Path "$pwd\report_$($SourceGPOSingle).xml" -Pattern "<Extension " -SimpleMatch | Where-Object {$_ -notlike "*RegistrySettings*"}
            If (($nonRegistry | Measure-Object).Count -gt 0) {
                Write-Warning "SOURCE GPO CONTAINS NON-REGISTRY SETTINGS FOR MANUAL COPY [$SourceGPOSingle]"
                Write-Warning ($nonRegistry -join "`r`n")
            }

        } Else {
            Write-Warning "SOURCE GPO DOES NOT EXIST [$SourceGPOSingle]"
        }

        $ProgressCounter++
    }

    Write-Progress -Activity "Copying GPO settings to: $DestinationGPO" -Completed -Status "Complete"

}

<#########################################################################sdg#>
BREAK

# Help
Help Get-GPLink -Full
Help Get-GPUnlinked -Full
Help Copy-GPRegistryValue -Full

# Copy one GPO registry settings into another
Copy-GPRegistryValue -Mode All -SourceGPO 'Client Settings' `
    -DestinationGPO 'New Merged GPO' -Verbose

# Copy one GPO registry settings into another, just user settings
Copy-GPRegistryValue -Mode User -SourceGPO 'Client Settings' `
    -DestinationGPO 'New Merged GPO' -Verbose

# Copy one GPO registry settings into another, just computer settings
Copy-GPRegistryValue -Mode Computer -SourceGPO 'Client Settings' `
    -DestinationGPO 'New Merged GPO' -Verbose

# Copy multiple GPO registry settings into another
Copy-GPRegistryValue -Mode All  -DestinationGPO "NewMergedGPO" `
    -SourceGPO "Firewall Policy", "Starter User", "Starter Computer" -Verbose

# Copy all GPOs linked to one OU registry settings into another
# Sort in reverse precedence order so that the highest precedence settings overwrite
# any potential settings conflicts in lower precedence policies.
$SourceGPOs = Get-GPLink -Path 'OU=SubTest,OU=Testing,DC=CohoVineyard,DC=com' |
    Sort-Object Precedence -Descending |
    Select-Object -ExpandProperty DisplayName
Copy-GPRegistryValue -Mode All -SourceGPO $SourceGPOs `
    -DestinationGPO "NewMergedGPO" -Verbose

# Log all GPO copy output (including verbose and warning)
# Requires PowerShell v3.0+
Copy-GPRegistryValue -Mode All -SourceGPO 'IE Test' `
    -DestinationGPO 'New Merged GPO' -Verbose *> GPOCopyLog.txt

# Disable all GPOs linked to an OU
Get-GPLink -Path 'OU=SubTest,OU=Testing,DC=CohoVineyard,DC=com' |
    ForEach-Object {
        Set-GPLink -Target $_.OUDN -GUID $_.GUID -LinkEnabled No -Confirm
    }

# Enable all GPOs linked to an OU
Get-GPLink -Path 'OU=SubTest,OU=Testing,DC=CohoVineyard,DC=com' |
    ForEach-Object {
        Set-GPLink -Target $_.OUDN -GUID $_.GUID -LinkEnabled Yes -Confirm
    }

# Quick link status of all GPOs
Get-GPUnlinked | Out-Gridview

# Just the unlinked GPOs
Get-GPUnlinked | Where-Object {!$_.Linked} | Out-GridView

# Detailed GP link status for all GPO with links
Get-GPLink | Out-GridView

# List of GPOs linked to a specific OU (or domain root)
Get-GPLink -Path 'OU=SubTest,OU=Testing,DC=CohoVineyard,DC=com' |
    Select-Object -ExpandProperty DisplayName

# List of OUs (or domain root) where a specific GPO is linked
Get-GPLink |
    Where-Object {$_.DisplayName -eq 'Script And Delegation Test'} |
    Select-Object -ExpandProperty OUDN
Function Invoke-Ping
{
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
Function Launch-AzurePortal { Invoke-Item "https://portal.azure.com/" -Credential (Get-Credential) }
Function Launch-ExchangeOnline { Invoke-Item "https://outlook.office365.com/ecp/" }
Function Launch-InternetExplorer { & 'C:\Program Files (x86)\Internet Explorer\iexplore.exe' "about:blank" }
Function Launch-Office365Admin { Invoke-Item "https://portal.office.com" -Credential (Get-Credential) }
Function Lock-Computer
{
	<#
		.DESCRIPTION
		Function to Lock your computer
		.SYNOPSIS
		Function to Lock your computer
	#>
	
$signature = @"
[DllImport("user32.dll", SetLastError = true)]
public static extern bool LockWorkStation();
"@

	$LockComputer = Add-Type -memberDefinition $signature -name "Win32LockWorkStation" -namespace Win32Functions -passthru
	$LockComputer::LockWorkStation() | Out-Null
}
Function MergeCSV {
  $Date = Get-Date -Format "d.MMM.yyyy"
  $path = "C:\LazyWinAdmin\Logs\Server-Apps\CSV\*"
  $csvs = Get-ChildItem $path -Include *.csv
  $y = $csvs.Count
  Write-Host "Detected the following CSV files: ($y)"
  foreach ($csv in $csvs) {
    Write-Host " "$csv.Name
  }
  $outputfilename = "Final Registry Results"
  Write-Host Creating: $outputfilename
  $excelapp = New-Object -ComObject Excel.Application
  $excelapp.SheetsInNewWorkbook = $csvs.Count
  $xlsx = $excelapp.Workbooks.Add()
  $sheet = 1
  foreach ($csv in $csvs) {
    $row = 1
    $column = 1
    $worksheet = $xlsx.Worksheets.Item($sheet)
    $worksheet.Name = $csv.Name
    $file = (Get-Content $csv)
    foreach ($line in $file) {
      $linecontents = $line -split ',(?!\s*\w+")'
      foreach ($cell in $linecontents) {
        $worksheet.Cells.Item($row,$column) = $cell
        $column++
      }
      $column = 1
      $row++
    }
    $sheet++
  }
  $output = "C:\LazyWinAdmin\Logs\Server-Apps\$Date\Results.Xlsx"
  $xlsx.SaveAs($output)
  $excelapp.Quit()
}
####################
# Static Functions #
####################
# Get-IPAddress
Function Get-IPAddress
{
	Get-NetIPAddress | ?{($_.interfacealias -notlike "*loopback*")`
 -and ($_.interfacealias -notlike "*vmware*")`
  -and ($_.interfacealias -notlike "*loopback*")`
   -and ($_.interfacealias -notlike "*bluetooth*")`
    -and ($_.interfacealias -notlike "*isatap*")} | ft
}
# Reload Profile
Function Reload-Profile {
    @(
        $Profile.AllUsersAllHosts,
        $Profile.AllUsersCurrentHost,
        $Profile.CurrentUserAllHosts,
        $Profile.CurrentUserCurrentHost
    ) | % {
        if(Test-Path $_){
            Write-Verbose "Running $_"
            . $_
        }
    }    
}
# End Get-IPAddress
# Begin RDP
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
# End RDP
# Begin Get-Lastboot
Function Get-LastBoot {
  <# 
  .SYNOPSIS 
  Retrieve last restart time for specified workstation(s) 

  .EXAMPLE 
  Get-LastBoot Computer123456 

  .EXAMPLE 
  Get-LastBoot 123456 
  #> 
    param([Parameter(Mandatory=$true)]
	[string[]] $ComputerName)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

$i=0
$j=0

    foreach ($Computer in $ComputerName) {

    Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

 
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{
            "Computer Name" = $Computer
            "Last Reboot"= $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}
# End Get-LastBoot
# Get-LoggedOnUser
Function global:Get-LoggedOnUser {
# PS version: 2.0 (tested Win7+)
# Written by: Yossi Sassi (yossis@protonmail.com) 
# Script version: 1.2
# Updated: August 14th, 2019

<# 
.SYNOPSIS

Gets current interactively logged-on users on all enabled domain computers, and check if they are a Direct member of Local Administrators group
(Not from Group membership (e.g. "Domain admins"), but were directly added to the local administrators group)

.DESCRIPTION

Gets currently logged-on users (interactive logins) on all computer accounts in the domain, and reports whether the logged-on user
is member of the local administrators group on that machine. This Function does not require any external module, all code provided as is in the Function.

.PARAMETER File

The name and location of the report file (Defaults to c:\LoggedOn.txt).

.PARAMETER ShowResultsToScreen

When specified, this switch shows the data collected in real time in the console, in addition to the log file.

.PARAMETER DoNotPingComputer

By Default - computers will first be pinged for 10ms timeout. If not responding, computer will be skipped. 
When specifying -DoNotPingComputer parameter, computer will be queried and tried access even if ping/ICMP echo response is blocked.
   
.EXAMPLE

PS C:\> Get-LoggedOnUser -File c:\temp\users-report.log
Sets the currently logged-on users report file to be saved at c:\temp\users-report.log.
Default is c:\LoggedOn.txt.

.EXAMPLE

PS C:\> Get-LoggedOnUser -ShowResultsToScreen
Shows the data collected in real time, onto the screen, in addition to the log file.

e.g.
LON-DC1	No User logged On interactively	False
LON-CL1	ADATUM\Administrator	True
LON-SVR1	ADATUM\adam	False
MSL1	ADATUM\yossis	False
The full report was saved to c:\LoggedOn.txt

.EXAMPLE

PS C:\> Import-Csv .\LoggedOn.txt -Delimiter "`t" | ft -AutoSize
Imports the CSV report file into Powershell, and lists the data in a table.

e.g.
HostName Logged-OnUserOrHostStatus       IsDirectLocalAdmin
-------- -------------------------       ------------------
LON-DC1  No User logged On interactively False   
LON-CL1  ADATUM\Administrator            True    
LON-SVR1 ADATUM\adam                     False   
MSL1     ADATUM\yossis                   False   

.EXAMPLE

PS C:\> $loggedOn = Import-Csv c:\LoggedOn.txt -Delimiter "`t"; $loggedOn | sort IsDirectLocalAdmin -Descending | ft -AutoSize
Gets the content of the report file into a variable, and outputs the results into a table, sorted by 'IsDirectLocalAdmin' property.

e.g.
HostName Logged-OnUserOrHostStatus       IsDirectLocalAdmin
-------- -------------------------       ------------------
LON-CL1  ADATUM\Administrator            True    
MSL1     ADATUM\yossis                   False   
LON-SVR1 ADATUM\adam                     False   
LON-DC1  No User logged On interactively False
#>
[cmdletbinding()]
param ([switch]$ShowResultsToScreen, 
[switch]$DoNotPingComputer,
[string]$File = "$ENV:TEMP\LoggedOn.txt"
 )

# Initialize
Write-Host "Initializing query. please wait...`n" -ForegroundColor cyan

# Check for number of computer accounts in the domain. If over 500, suggest potential alternatives
# Get all Enabled computer accounts 
$Searcher = New-Object DirectoryServices.DirectorySearcher([ADSI]"")
$Searcher.Filter = "(&(objectClass=computer)(!userAccountControl:1.2.840.113556.1.4.803:=2))"
$Searcher.PageSize = 50000 # by default, 1000 are returned for adsiSearcher. this script will handle up to 50K acccounts.
$Computers = ($Searcher.Findall())

if ($Computers.count -gt 500) {
$PromptText = "You have $($computers.count) enabled computer accounts in domain $env:USERDNSDOMAIN.`nAre you sure you want to proceed?`nNote: Running this script over the network could take a while, and in large AD networks you might prefer running it locally using SCCM, PSRemoting etc."
$PromptTitle = "Get-LoggedOnUser"
$Options = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$Options.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
$Options.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
$Choice = $host.ui.PromptForChoice($PromptTitle,$PromptText,$Options,0)
If ($Choice -eq 1) {break}}

# If OK - continue with the script
# Get the current Error Action Preference
$CurrentEAP = $ErrorActionPreference
# Set script not to alert for errors
$ErrorActionPreference = "silentlycontinue"
$report = @()
$report += "HostName`tLogged-OnUserOrHostStatus`tIsDirectLocalAdmin"
$OfflineComputers = @()

# If not responding to Ping - by default, host will be skipped. 
# NOTE: Default timeout for ping is 10ms - you can change it in the following Function below
filter Invoke-Ping {(New-Object System.Net.NetworkInformation.Ping).Send($_,10)}

foreach ($comp in $Computers)
    { 
    # Check if computer needs to be Pinged first or not, and if Yes - see if responds to ping    
     switch ($DoNotPingComputer)
     {
     $false {$ProceedToCheck = ($Comp.Properties.dnshostname | Invoke-Ping).status -eq "Success"}
     $true {$ProceedToCheck = $true}
    }
     
     if ($ProceedToCheck) {   
     $user = gwmi win32_computersystem -ComputerName $Comp.Properties.dnshostname | select -ExpandProperty username1
# If wmi query returned empty results - try querying with QUSER for active console session 
if ($user -eq $null) {
$user = quser /SERVER:$($Comp.Properties.dnshostname) | select-string active | % {$_.toString().split(" ")[1].Trim()}
} 

# Check if logged on user is a Direct member of Local Administrators group
     if ($user -eq $null) {$user = "No User logged On interactively"} 
        else # Check if local admin
        # Note: locally can be checkd as- [Security.Principal.WindowsIdentity]::GetCurrent().IsInRole([Security.PrincipaltInRole] "Administrator")        
        {
        $group = [ADSI]"WinNT://$($Comp.Properties.dnshostname)/administrators,group"
        $member=@($group.psbase.invoke("Members"))      
        $usersInGroup = $member | ForEach-Object {([ADSI]$_).InvokeGet("Name")} 
        foreach ($GroupEntry in $usersInGroup) 
            {if ($GroupEntry -eq $user) {$AdminRole = $true}}
        }
     if ($AdminRole -ne $true -and $user -ne $null) {$AdminRole = $false} # if not admin, set to false     
     if ($ShowResultsToScreen) {write-host "$($Comp.Properties.dnshostname)`t$user`t$AdminRole"}
     $report += "$($Comp.Properties.dnshostname)`t$user`t$AdminRole"
     $user = $null
     $adminRole = $null
     $group = $null
     $member = $null
     $usersInGroup = $null
     } 
     else 
     # computer didn't respond to ping     
      {$report += $($Comp.Properties.dnshostname) + "`tdidn't respond to ping - possibly Offile or Firewall issue"; $OfflineComputers += $($comp.properties.name)
      if ($ShowResultsToScreen) {Write-Warning "$($Comp.Properties.dnshostname)`tdidn't respond to ping - possibly  Offile or Port issue"}
      }
    }
$report | Out-File $File 

# Wrap up
Write-Host "`nCompleted checking $($Computers.Count) hosts.`n" -ForegroundColor Green

# check for offline computers, if encountered
If ($OfflineComputers -ne $null) # If there were offline / Non-responsive computers
{ $OfflineComputers | Out-File "$ENV:Temp\NonRespondingComputers.txt"
  Write-Warning "Total of $($OfflineComputers.count) computers didn't respond to Ping.`nNon-Responding computers where saved into $($ENV:Temp)\NonRespondingComputers.txt." 
 }

Write-Host "The full report was saved to $File" -ForegroundColor Cyan
# Set back the system's current Error Action Preference
$ErrorActionPreference = $CurrentEAP
}
#End Get-LoggedOnUser

# Get-HotFixes
Function Get-HotFixes {
  <# 
  .SYNOPSIS 
  Grabs all processes on specified workstation(s).

  .EXAMPLE 
  Get-HotFixes Computer123456 

  .EXAMPLE 
  Get-HotFixes 123456 
  #> 
param (
    [Parameter(ValueFromPipeline=$true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [string]$NameRegex = '')

if(($computername.length -eq 6)) {
    [int32] $dummy_output = $null;

    if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        $computername = "Computer" + $computername.Replace("Computer","")
    }	
}

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

Function HotFix {

$i=0
$j=0

    foreach ($computer in $ComputerArray) {

        Write-Progress -Activity "Retrieving HotFix Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

        Get-HotFix -Computername $computer 
    }    
}

foreach ($computer in $ComputerName) {	     
    If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
        $ComputerArray += $Computer
    }	
}

$HotFix = HotFix
$DocPath = [environment]::getfolderpath("mydocuments") + "\HotFix-Report.csv"

    		Switch ($CheckBox.IsChecked){
    		    $true { $HotFix | Export-Csv $DocPath -NoTypeInformation -Force; }
    		    default { $HotFix | Out-GridView -Title "HotFix Report"; }
    		}

	if($CheckBox.IsChecked -eq $true){
	    Try { 
		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {
		 #Do Nothing 
	    }
	}
	
	else{
	    Try {
	        $listBox.Items.Add("$stamp HotFixes output processed!`n")
	    } 
	    Catch {
	        #Do Nothing 
	    }
	}
}
# End Get-HotFixes
# Begin Get-GPRemote

Function Get-GPRemote {
  <# 
  .SYNOPSIS 
  Open Group Policy for specified workstation(s) 

  .EXAMPLE 
  Get-GPRemote Computer123456 

  .EXAMPLE 
  Get-GPRemote 123456 
  #> 
param(
[Parameter(Mandatory=$true)]
[string[]] $ComputerName)

if (($computername.length -eq 6)) {
    [int32] $dummy_output = $null;

    if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
       	$computername = "Computer" + $computername.Replace("Computer","")}	
}

$i=0
$j=0

foreach ($Computer in $ComputerName) {

    Write-Progress -Activity "Opening Remote Group Policy..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	#Opens (Remote) Group Policy for specified workstation
	gpedit.msc /gpcomputer: $Computer
    
	}
}
# End Get-GPRemote

# Begin CheckProcess
Function CheckProcess {
  <# 
  .SYNOPSIS 
  Grabs all processes on specified workstation(s).

  .EXAMPLE 
  CheckProcess Computer123456 

  .EXAMPLE 
  CheckProcess 123456 
  #> 
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = '')
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

Function ChkProcess {

$i=0
$j=0

    foreach ($computer in $ComputerArray) {

        Write-Progress -Activity "Retrieving System Processes..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

        $getProcess = Get-Process -ComputerName $computer

        foreach ($Process in $getProcess) {
                
             [pscustomobject]@{
		"Computer Name" = $computer
                "Process Name" = $Process.ProcessName
                PID = '{0:f0}' -f $Process.ID
                Company = $Process.Company
                "CPU(s)" = $Process.CPU
                Description = $Process.Description
             }           
         }
     } 
}
	
foreach ($computer in $ComputerName) {	     
    If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
        $ComputerArray += $Computer
    }	
}
	$chkProcess = ChkProcess | Sort "Computer Name" | Select "Computer Name","Process Name", PID, Company, "CPU(s)", Description
    	$DocPath = [environment]::getfolderpath("mydocuments") + "\Process-Report.csv"

    		Switch ($CheckBox.IsChecked){
    		    $true { $chkProcess | Export-Csv $DocPath -NoTypeInformation -Force; }
    		    default { $chkProcess | Out-GridView -Title "Processes";  }
    		}

	if($CheckBox.IsChecked -eq $true){
	    Try { 
		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {
		 #Do Nothing 
	    }
	}
	
	else{
	    Try {
	        $listBox.Items.Add("$stamp Check Process output processed!`n")
	    } 
	    Catch {
	        #Do Nothing 
	    }
	}
    
}
# End CheckProcess
# Begin Whois

Function WhoIs
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
 {
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
# End WhoIs
# Begin Get-NetworkStatistics
Function Get-NetworkStatistics
<#
.SYNOPSIS
PowerShell version of netstat
.EXAMPLE
Get-NetworkStatistics
.EXAMPLE
Get-NetworkStatistics | where-object {$_.State -eq "LISTENING"} | Format-Table
#>
{ 
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
# End Get-NetworkStatistics
# Begin Update-SysInternals
Function Update-Sysinternals
<#
.Synopsis
   Download the latest sysinternals tools
.DESCRIPTION
   Downloads the latest sysinternals tools from https://live.sysinternals.com/ to a specified directory
   The Function downloads all .exe and .chm files available
.EXAMPLE
   Update-Sysinternals -Path C:\sysinternals
   Downloads the sysinternals tools to the directory C:\sysinternals
.EXAMPLE
   Update-Sysinternals -Path C:\Users\Matt\OneDrive\Tools\sysinternals
   Downloads the sysinternals tools to a user's OneDrive
#>
 {
    [CmdletBinding()]
    param (
        # Path to the directory were sysinternals tools will be downloaded to 
        [Parameter(Mandatory=$true)]      
        [string]
        $Path 
    )
    
    begin {
            if (-not (Test-Path -Path $Path)){
            Throw "The Path $_ does not exist"
        } else {
            $true
        }
        
            $uri = 'https://live.sysinternals.com/'
            $sysToolsPage = Invoke-WebRequest -Uri $uri
            
    }
    
    process {
        # create dir if it doesn't exist    
       
        Set-Location -Path $Path

        $sysTools = $sysToolsPage.Links.innerHTML | Where-Object -FilterScript {$_ -like "*.exe" -or $_ -like "*.chm"} 

        foreach ($sysTool in $sysTools){
            Invoke-WebRequest -Uri "$uri/$sysTool" -OutFile $sysTool
        }
    } #process
}
# End Update-SysInternals
# Begin Get-ADGPOReplication
Function Get-ADGPOReplication
{
	<#
	.SYNOPSIS
		This Function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.DESCRIPTION
		This Function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.PARAMETER GPOName
		Specify the name of the GPO
	.PARAMETER All
		Specify that you want to retrieve all the GPO (slow if you have a lot of Domain Controllers)
	.EXAMPLE
		Get-ADGPOReplication -GPOName "Default Domain Policy"
	.EXAMPLE
		Get-ADGPOReplication -All
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		lazywinadmin.com
	
		VERSION HISTORY
		1.0 2014.09.22 	Initial version
						Adding some more Error Handling
						Fix some typo
	#>
	

	[CmdletBinding()]
	PARAM (
		[parameter(Mandatory = $True, ParameterSetName = "One")]
		[String[]]$GPOName,
		[parameter(Mandatory = $True, ParameterSetName = "All")]
		[Switch]$All
	)
Remove-Module Carbon
	BEGIN
	{
		TRY
		{
			if (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction Stop -ErrorVariable ErrorBeginIpmoAD }
			if (-not (Get-Module -Name GroupPolicy)) { Import-Module -Name GroupPolicy -ErrorAction Stop -ErrorVariable ErrorBeginIpmoGP }
		}
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			IF ($ErrorBeginIpmoAD) { Write-Warning -Message "[BEGIN] Error while Importing the module Active Directory" }
			IF ($ErrorBeginIpmoGP) { Write-Warning -Message "[BEGIN] Error while Importing the module Group Policy" }
			Write-Warning -Message "[BEGIN] $($Error[0].exception.message)"
		}
	}
	PROCESS
	{
		FOREACH ($DomainController in ((Get-ADDomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetDC -filter *).hostname))
		{
			TRY
			{
				IF ($psBoundParameters['GPOName'])
				{
					Foreach ($GPOItem in $GPOName)
					{
						$GPO = Get-GPO -Name $GPOItem -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPO
						
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPOItem
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}#Foreach ($GPOItem in $GPOName)
				}#IF ($psBoundParameters['GPOName'])
				IF ($psBoundParameters['All'])
				{
					$GPOList = Get-GPO -All -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPOAll
					
					foreach ($GPO in $GPOList)
					{
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPO.DisplayName
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}
				}#IF ($psBoundParameters['All'])
			}#TRY
			CATCH
			{
				Write-Warning -Message "[PROCESS] Something wrong happened"
				IF ($ErrorProcessGetDC) { Write-Warning -Message "[PROCESS] Error while running retrieving Domain Controllers with Get-ADDomainController" }
				IF ($ErrorProcessGetGPO) { Write-Warning -Message "[PROCESS] Error while running Get-GPO" }
				IF ($ErrorProcessGetGPOAll) { Write-Warning -Message "[PROCESS] Error while running Get-GPO -All" }
				Write-Warning -Message "[PROCESS] $($Error[0].exception.message)"
			}
		}#FOREACH
	}#PROCESS
}
# End Get-ADGPOReplication
# Begin Get-LocalAdmin 
Function Get-LocalAdmin { 
param ($ComputerName) 
 
$admins = Gwmi win32_groupuser –computer $ComputerName  
$admins = $admins |? {$_.groupcomponent –like '*"Administrators"'} 
 
$admins |% { 
$_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul 
$matches[1].trim('"') + “\” + $matches[2].trim('"') 
} 
}
### End Get-LocalAdmin
####################
# Static Functions #
####################
# Touch
Function touch { $args | foreach-object {write-host > $_} }
# Notepad++
Function NPP { Start-Process -FilePath "${Env:ProgramFiles(x86)}\Notepad++\Notepad++.exe" }#-ArgumentList $args }
# Find File
Function findfile($name) {
	ls -recurse -filter "*${name}*" -ErrorAction SilentlyContinue | foreach {
		$place_path = $_.directory
		echo "${place_path}\${_}"
	}
}
# RM -RF
Function rm-rf($item) { Remove-Item $item -Recurse -Force }
# SUDO
Function sudo(){
	Invoke-Elevated @args
}
# SED
Function PSsed($file, $find, $replace){
	(Get-Content $file).replace("$find", $replace) | Set-Content $file
}
# SED-Recursive
Function PSsed-recursive($filePattern, $find, $replace) {
	$files = ls . "$filePattern" -rec # -Exclude
	foreach ($file in $files) {
		(Get-Content $file.PSPath) |
		Foreach-Object { $_ -replace "$find", "$replace" } |
		Set-Content $file.PSPath
	}
}
# PSGrep
Function PSgrep {

    [CmdletBinding()]
    Param(
    
        # source file to grep
        [Parameter(Mandatory=$true)]
        [string]$SourceFileName, 

        # string to search for
        [Parameter(Mandatory=$true)]
        [string]$SearchStrings,

        # do we write to file
        [Parameter()]
        [string]$OutputFile
    )

        # break the comma separated strings up
        $Strings = @()
        $Strings = $SearchStrings.split(',')
        $count = 0

        # write-host $Strings

        $Content = Get-Content $SourceFileName
        
        $Content | ForEach-Object { 
            foreach ($String in $Strings) {
                # $String
                if($_ -match $String){
                    $count ++
                    if (!($OutputFile)) {
                        write-host $_
                    } else {
                        $_ | Out-File -FilePath ".\$($OutputFile)" -Append -Force
                }

            }

        }

    }

    Write-Host "$($Count) matches found"
}
# End PSgrep
# Which
Function which($name) {
	Get-Command $name | Select-Object -ExpandProperty Definition
}
# Cut
Function cut(){
	foreach ($part in $input) {
		$line = $part.ToString();
		$MaxLength = [System.Math]::Min(200, $line.Length)
		$line.subString(0, $MaxLength)
	}
}
# Search Text Files
Function Search-AllTextFiles {
    param(
        [parameter(Mandatory=$true,position=0)]$Pattern, 
        [switch]$CaseSensitive,
        [switch]$SimpleMatch
    );

    Get-ChildItem . * -Recurse -Exclude ('*.dll','*.pdf','*.pdb','*.zip','*.exe','*.jpg','*.gif','*.png','*.ico','*.svg','*.bmp','*.psd','*.cache','*.doc','*.docx','*.xls','*.xlsx','*.dat','*.mdf','*.nupkg','*.snk','*.ttf','*.eot','*.woff','*.tdf','*.gen','*.cfs','*.map','*.min.js','*.data') | Select-String -Pattern:$pattern -SimpleMatch:$SimpleMatch -CaseSensitive:$CaseSensitive
}
# Add to Zip
Function AddTo-7zip($zipFileName) {
    BEGIN {
        #$7zip = "$($env:ProgramFiles)\7-zip\7z.exe"
        $7zip = Find-Program "\7-zip\7z.exe"
		if(!([System.IO.File]::Exists($7zip))){
			throw "7zip not found";
		}
    }
    PROCESS {
        & $7zip a -tzip $zipFileName $_
    }
    END {
    }
}
## End Add to Zip

# Connect to Exchange
Function GoGo-PSExch {

    param(
        [Parameter( Mandatory=$false)]
        [string]$URL="MWTEXCH01"
    )
    
    #$Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

    $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos #-Credential #$Credentials

    Import-PSSession $ExOPSession -AllowClobber

}
## End Connect to Exchange

## Connect to VMware VSphere
Function GoGo-VSphere {

Connect-VIServer -Server 10.20.1.9
}

## End Connect to VMware VSphere

## Out-File in UTF8 NonBom
Function Out-FileUtf8NoBom {

  [CmdletBinding()]
  param(
    [Parameter(Mandatory, Position=0)] [string] $LiteralPath,
    [switch] $Append,
    [switch] $NoClobber,
    [AllowNull()] [int] $Width,
    [Parameter(ValueFromPipeline)] $InputObject
  )

  

  # Make sure that the .NET framework sees the same working dir. as PS
  # and resolve the input path to a full path.
  [System.IO.Directory]::SetCurrentDirectory($PWD) # Caveat: .NET Core doesn't support [Environment]::CurrentDirectory
  $LiteralPath = [IO.Path]::GetFullPath($LiteralPath)

  # If -NoClobber was specified, throw an exception if the target file already
  # exists.
  if ($NoClobber -and (Test-Path $LiteralPath)) {
    Throw [IO.IOException] "The file '$LiteralPath' already exists."
  }

  # Create a StreamWriter object.
  # Note that we take advantage of the fact that the StreamWriter class by default:
  # - uses UTF-8 encoding
  # - without a BOM.
  $sw = New-Object IO.StreamWriter $LiteralPath, $Append

  $htOutStringArgs = @{}
  if ($Width) {
    $htOutStringArgs += @{ Width = $Width }
  }

  # Note: By not using begin / process / end blocks, we're effectively running
  #       in the end block, which means that all pipeline input has already
  #       been collected in automatic variable $Input.
  #       We must use this approach, because using | Out-String individually
  #       in each iteration of a process block would format each input object
  #       with an indvidual header.
  try {
    $Input | Out-String -Stream @htOutStringArgs | % { $sw.WriteLine($_) }
  } finally {
    $sw.Dispose()
  }

}
## End Out-File in UTF8 NonBom

## Invoke VBScript
Function Invoke-VBScript {
    <#
    .Synopsis
       Run VBScript from PowerShell
    .DESCRIPTION
       Used to invoke VBScript from PowerShell

       Will run the VBScript in a separate job using cscript.exe
    .PARAMETER Path
       Path to VBScript.
       Accepts relative or absolute path.
    .PARAMETER Argument
       Arguments to pass to VBScript
    .PARAMETER Wait
       Wait for VBScript to finish   
    .EXAMPLE
       Invoke-VBScript -Path '.\VBScript1.vbs' -Arguments '"MyFirstArgument"', '"MySecondArgument"' -Wait
       Run VBScript1.vbs using cscript and wait for the script to complete.
       Displays progressbar while waiting.
       Returns script output as single string.
    .EXAMPLE
       '.\VBScript1.vbs', '.\VBScript2.vbs' | Invoke-VBScript -Arguments '"MyArgument"'
       Starts both VBScript1.vbs and VBScript2.vbs in separate jobs simultaneously.
       Both scripts will be run using the same arguments.
       Returns job items.
    .EXAMPLE
       [PSCustomObject]@{Path='.\VBScript1.vbs';Arguments='"Script1"'},[PSCustomObject]@{Path='.\VBScript2.vbs';Arguments='"Script2"'} | Invoke-VBScript -Wait -Verbose
       Runs two scripts after each other, waiting to one to complete
       before starting next.
       Each script will run with different parameters.
       Displays progressbar while waiting.
       Returns script output in one single string per script.
    .NOTES
       Written by Simon Wåhlin
       http://blog.simonw.se
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='None',PositionalBinding=$false)]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [ValidateScript({if(Test-Path $_){$true}else{Throw "Could not find script: [$_]"}})]
        [String]
        $Path,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Args')]
        [String[]]
        $Argument,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch]
        $Wait
    )
    Begin
    {
        Write-Verbose -Message 'Locating cscript.exe'
        $cscriptpath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\cscript.exe'
        if(-Not(Test-Path -Path $cscriptpath))
        {
            Throw 'cscript.exe not found.'
        }
        Write-Verbose -Message ('cscript.exe found in: {0}' -f $cscriptpath)
    }
    Process
    {
        Try
        {
            $ResolvedPath = Resolve-Path -Path $Path
            Write-Verbose -Message ('Processing script: {0}' -f $ResolvedPath)
            if($PSBoundParameters.ContainsKey('Argument'))
            {
                $ScriptBlock = [scriptblock]::Create(('& "{0}" "{1}" "{2}"' -f $cscriptpath, $ResolvedPath,($Argument -join '" "')))
            }
            else
            {
                $ScriptBlock = $ScriptBlock = [scriptblock]::Create(('& "{0}" "{1}"' -f $cscriptpath, $ResolvedPath))
            }
            Write-Verbose -Message 'Starting script'
            if($PSCmdlet.ShouldProcess($ResolvedPath,'Invoke script'))
            {
                $Job = Start-Job -ScriptBlock $ScriptBlock
                if($Wait)
                {
                    $Activity = 'Waiting for script to complete: {0}' -f $ResolvedPath
                    Write-Progress -Activity $Activity -Id 1
                    $i = 1
                    While($Job.State -eq 'Running')
                    {
                        $WaitTime = (Get-Date) - $Job.PSBeginTime
                        Write-Progress -Activity $Activity -Status "Waited for $($WaitTime.TotalSeconds -as [int]) seconds." -Id 1 -PercentComplete ($i%100)
                        Start-Sleep -Seconds 1
                        $i++
                    }
                    Write-Progress -Activity $Activity -Status 'Waiting' -Id 1 -Completed
                    $Result = Foreach($JobInstance in ($Job,$Job.ChildJobs))
                    {
                        if($JobInstance.Error -ne $null)
                        {
                            Throw $JobInstance.Error.Exception.Message
                        }
                        else
                        {
                            $JobInstance.Output
                        }
                    }
                    Write-Output -InputObject ($Result -join "`n")
                    Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
                }
                else
                {
                    Write-Output -InputObject $Job
                }
            }
            Write-Verbose -Message 'Finished processing script'
        }
        Catch
        {
            Throw
        }
    }
}
## End Invoke VBScript
## Function Get-MOTD
Function Get-MOTD {

<#
.NAME
    Get-MOTD
.SYNOPSIS
    Displays system information to a host.
.DESCRIPTION
    The Get-MOTD cmdlet is a system information tool written in PowerShell. 
.EXAMPLE
#>


  [CmdletBinding()]
	
  Param(
    [Parameter(Position=0,Mandatory=$false)]
	[ValidateNotNullOrEmpty()]
    [string[]]$ComputerName
    ,
    [Parameter(Position=1,Mandatory=$false)]
    [PSCredential]
    [System.Management.Automation.CredentialAttribute()]$Credential
  )

  Begin {
	
        If (-Not $ComputerName) {
            $RemoteSession = $null
        }
        #Define ScriptBlock for data collection
        $ScriptBlock = {
            $Operating_System = Get-CimInstance -ClassName Win32_OperatingSystem
            $Logical_Disk = Get-CimInstance -ClassName Win32_LogicalDisk |
            Where-Object -Property DeviceID -eq $Operating_System.SystemDrive
			Try {
				$PCLi = Get-PowerCLIVersion
				$PCLiVer = ' | PowerCLi ' + [string]$PCLi.Major + '.' + [string]$PCLi.Minor + '.' + [string]$PCLi.Revision + '.' + [string]$PCLi.Build
			} Catch {$PCLiVer = ''}
			If ($DomainName = ([System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()).DomainName) {$DomainName = '.' + $DomainName}
			
            [pscustomobject]@{
                Operating_System = $Operating_System
                Processor = Get-CimInstance -ClassName Win32_Processor
                Process_Count = (Get-Process).Count
                Shell_Info = ("{0}.{1}" -f $PSVersionTable.PSVersion.Major,$PSVersionTable.PSVersion.Minor) + $PCLiVer
                Logical_Disk = $Logical_Disk
            }
        }
  } #End Begin

  Process {
	
        If ($ComputerName) {
            If ("$ComputerName" -ne "$env:ComputerName") {
                # Build Hash to be used for passing parameters to 
                # New-PSSession commandlet
                $PSSessionParams = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }

                # Add optional parameters to hash
                If ($Credential) {
                    $PSSessionParams.Add('Credential', $Credential)
                }

                # Create remote powershell session   
                Try {
                    $RemoteSession = New-PSSession @PSSessionParams
                }
                Catch {
                    Throw $_.Exception.Message
                }
            } Else { 
                $RemoteSession = $null
            }
        }
        
        # Build Hash to be used for passing parameters to 
        # Invoke-Command commandlet
        $CommandParams = @{
            ScriptBlock = $ScriptBlock
            ErrorAction = 'Stop'
        }
        
        # Add optional parameters to hash
        If ($RemoteSession) {
            $CommandParams.Add('Session', $RemoteSession)
        }
               
        # Run ScriptBlock    
        Try {
            $ReturnedValues = Invoke-Command @CommandParams
        }
        Catch {
            If ($RemoteSession) {
            	Remove-PSSession $RemoteSession
            }
            Throw $_.Exception.Message
        }

        # Assign variables
        #Import-Module MS-Module
        $Date = Get-Date
        $OS_Name = $ReturnedValues.Operating_System.Caption + ' [Installed: ' + ([datetime]$ReturnedValues.Operating_System.InstallDate).ToString('dd-MMM-yyyy') + ']'
        $Computer_Name = $ReturnedValues.Operating_System.CSName
		If ($DomainName) {$Computer_Name = $Computer_Name + $DomainName.ToUpper()}
        $Kernel_Info = $ReturnedValues.Operating_System.Version + ' [' + $ReturnedValues.Operating_System.OSArchitecture + ']'
        $Process_Count = $ReturnedValues.Process_Count
        $Uptime = "$(($Uptime = $Date - $($ReturnedValues.Operating_System.LastBootUpTime)).Days) days, $($Uptime.Hours) hours, $($Uptime.Minutes) minutes"
        $Shell_Info = $ReturnedValues.Shell_Info
        $CPU_Info = $ReturnedValues.Processor.Name -replace '\(C\)', '' -replace '\(R\)', '' -replace '\(TM\)', '' -replace 'CPU', '' -replace '\s+', ' '
        $Current_Load = $ReturnedValues.Processor.LoadPercentage    
        $Memory_Size = "{0} MB/{1} MB " -f (([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))-
        ([math]::round($ReturnedValues.Operating_System.FreePhysicalMemory/1KB))),([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))
		$Disk_Size = "{0} GB/{1} GB" -f (([math]::round($ReturnedValues.Logical_Disk.Size/1GB)-
        [math]::round($ReturnedValues.Logical_Disk.FreeSpace/1GB))),([math]::round($ReturnedValues.Logical_Disk.Size/1GB))

        # Write to the Console
        Write-Host -Object ("")
        Write-Host -Object ("")
        Write-Host -Object ("         ,.=:^!^!t3Z3z.,                  ") -ForegroundColor Red
        Write-Host -Object ("        :tt:::tt333EE3                    ") -ForegroundColor Red
        Write-Host -Object ("        Et:::ztt33EEE ") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @Ee.,      ..,     $($Date.ToString('dd-MMM-yyyy HH:mm:ss'))") -ForegroundColor Green
        Write-Host -Object ("       ;tt:::tt333EE7") -NoNewline -ForegroundColor Red
        Write-Host -Object (" ;EEEEEEttttt33#     ") -ForegroundColor Green
        Write-Host -Object ("      :Et:::zt333EEQ.") -NoNewline -ForegroundColor Red
        Write-Host -Object (" SEEEEEttttt33QL     ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("User: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$env:USERDOMAIN\$env:UserName") -ForegroundColor Cyan
        Write-Host -Object ("      it::::tt333EEF") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @EEEEEEttttt33F      ") -NoNewline -ForeGroundColor Green
        Write-Host -Object ("Hostname: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Computer_Name") -ForegroundColor Cyan
        Write-Host -Object ("     ;3=*^``````'*4EEV") -NoNewline -ForegroundColor Red
        Write-Host -Object (" :EEEEEEttttt33@.      ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("OS: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$OS_Name") -ForegroundColor Cyan
        Write-Host -Object ("     ,.=::::it=., ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("``") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @EEEEEEtttz33QF       ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("Kernel: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("NT ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("$Kernel_Info") -ForegroundColor Cyan
        Write-Host -Object ("    ;::::::::zt33) ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("  '4EEEtttji3P*        ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("Uptime: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Uptime") -ForegroundColor Cyan
        Write-Host -Object ("   :t::::::::tt33.") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (":Z3z.. ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object (" ````") -NoNewline -ForegroundColor Green
        Write-Host -Object (" ,..g.        ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Shell: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("PowerShell $Shell_Info") -ForegroundColor Cyan
        Write-Host -Object ("   i::::::::zt33F") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" AEEEtttt::::ztF         ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("CPU: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$CPU_Info") -ForegroundColor Cyan
        Write-Host -Object ("  ;:::::::::t33V") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" ;EEEttttt::::t3          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Processes: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Process_Count") -ForegroundColor Cyan
        Write-Host -Object ("  E::::::::zt33L") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" @EEEtttt::::z3F          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Current Load: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Current_Load") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("%") -ForegroundColor Cyan
        Write-Host -Object (" {3=*^``````'*4E3)") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" ;EEEtttt:::::tZ``          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Memory: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Memory_Size`t") -ForegroundColor Cyan -NoNewline
		New-PercentageBar -DrawBar -Value (([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))-([math]::round($ReturnedValues.Operating_System.FreePhysicalMemory/1KB))) -MaxValue ([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB)); "`r"
        Write-Host -Object ("             ``") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" :EEEEtttt::::z7            ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("System Volume: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Disk_Size`t") -ForegroundColor Cyan -NoNewline
		New-PercentageBar -DrawBar -Value (([math]::round($ReturnedValues.Logical_Disk.Size/1GB)-[math]::round($ReturnedValues.Logical_Disk.FreeSpace/1GB))) -MaxValue ([math]::round($ReturnedValues.Logical_Disk.Size/1GB)); "`r"
        Write-Host -Object ("                 'VEzjt:;;z>*``           ") -ForegroundColor Yellow
        Write-Host -Object ("                      ````                  ") -ForegroundColor Yellow
        Write-Host -Object ("")
  } #End Process

  End {
        If ($RemoteSession) {
            Remove-PSSession $RemoteSession
        }
  }
} #End Function Get-MOTD

## Change Attributes
Function Get-FileAttribute{
    param($file,$attribute)
    $val = [System.IO.FileAttributes]$attribute;
    if((gci $file -force).Attributes -band $val -eq $val){$true;} else { $false; }
} 


Function Set-FileAttribute{
    param($file,$attribute)
    $file =(gci $file -force);
    $file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]$attribute).value__;
    if($?){$true;} else {$false;}
} 

## End Change Attributes

## Remote Group Policy
Function GPR {
<# 
.SYNOPSIS 
    Open Group Policy for specified workstation(s) 

.EXAMPLE 
    GPR Computer123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [String[]]$ComputerName,

    $i=0,
    $j=0
)

    foreach ($Computer in $ComputerName) {

        Write-Progress -Activity "Opening Remote Group Policy..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

        #Opens (Remote) Group Policy for specified workstation
        GPedit.msc /gpcomputer: $Computer
    }
}#End GPR

## Begin Lastboot

Function LastBoot {
<# 
.SYNOPSIS 
    Retrieve last restart time for specified workstation(s) 

.EXAMPLE 
    LastBoot Computer123456 

.EXAMPLE 
    LastBoot 123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [String[]]$ComputerName,

    $i=0,
    $j=0
)

    foreach($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)
 
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{

            ComputerName = $Computer
            LastReboot = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}#End LastBoot

#Begin SYSinfo
Function SYSinfo {
<# 
.SYNOPSIS 
  Retrieve basic system information for specified workstation(s) 

.EXAMPLE 
  SYS Computer123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [string[]] $ComputerName,
    
    $i=0,
    $j=0
)

$Stamp = (Get-Date -Format G) + ":"

    Function Systeminformation {
	
        foreach ($Computer in $ComputerName) {

            if(!([String]::IsNullOrWhiteSpace($Computer))) {

                if(Test-Connection -Quiet -Count 1 -Computer $Computer) {

                    Write-Progress -Activity "Getting Sytem Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	                Start-Job -ScriptBlock { param($Computer) 

	                    #Gather specified workstation information; CimInstance only works on 64-bit
	                    $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer
	                    $computerBIOS = Get-CimInstance CIM_BIOSElement -Computer $Computer
	                    $computerOS = Get-CimInstance CIM_OperatingSystem -Computer $Computer
	                    $computerCPU = Get-CimInstance CIM_Processor -Computer $Computer
	                    $computerHDD = Get-CimInstance Win32_LogicalDisk -Computer $Computer -Filter "DeviceID = 'C:'"
    
                        [PSCustomObject] @{

                            ComputerName = $computerSystem.Name
                            LastReboot = $computerOS.LastBootUpTime
                            OperatingSystem = $computerOS.OSArchitecture + " " + $computerOS.caption
                            Model = $computerSystem.Model
                            RAM = "{0:N2}" -f [int]($computerSystem.TotalPhysicalMemory/1GB) + "GB"
                            DiskCapacity = "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
                            TotalDiskSpace = "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
                            CurrentUser = $computerSystem.UserName
                        }
                    } -ArgumentList $Computer
                }

                else {

                    Start-Job -ScriptBlock { param($Computer)  
                     
                        [PSCustomObject] @{

                            ComputerName=$Computer
                            LastReboot="Unable to PING."
                            OperatingSystem="$Null"
                            Model="$Null"
                            RAM="$Null"
                            DiskCapacity="$Null"
                            TotalDiskSpace="$Null"
                            CurrentUser="$Null"
                        }
                    } -ArgumentList $Computer                       
                }
            }

            else {
                 
                Start-Job -ScriptBlock { param($Computer)  
                     
                    [PSCustomObject] @{

                        ComputerName = "Value is null."
                        LastReboot = "$Null"
                        OperatingSystem = "$Null"
                        Model = "$Null"
                        RAM = "$Null"
                        DiskCapacity = "$Null"
                        TotalDiskSpace = "$Null"
                        CurrentUser = "$Null"
                    }
                } -ArgumentList $Computer
            }
        } 
    }

    $SystemInformation = SystemInformation | Receive-Job -Wait | Select ComputerName, CurrentUser, OperatingSystem, Model, RAM, DiskCapacity, TotalDiskSpace, LastReboot
    $DocPath = [environment]::getfolderpath("mydocuments") + "\SystemInformation-Report.csv"

	Switch($CheckBox.IsChecked) {

		$true { 
            
            $SystemInformation | Export-Csv $DocPath -NoTypeInformation -Force 
        }

		default { 
            
            $SystemInformation | Out-GridView -Title "System Information"
        }
    }

	if($CheckBox.IsChecked -eq $true) {

	    Try { 

		    $listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {

		    #Do Nothing 
	    }
	}
	
	else {

	    Try {

	        $listBox.Items.Add("$stamp System Information output processed!`n")
	    } 

	    Catch {

	        #Do Nothing 
	    }
	}
}#End SYSinfo

#Begin NetMessage

Function NetMSG {
<# 
.SYNOPSIS 
    Generate a pop-up window on specified workstation(s) with desired message 

.EXAMPLE 
    NetMSG Computer123456 
#> 
	
param(

    [Parameter(Mandatory=$true)]
    [String[]] $ComputerName,

    [Parameter(Mandatory=$true,HelpMessage='Enter desired message')]
    [String]$MyMessage,

    [String]$User = [Environment]::UserName,

    [String]$UserJob = (Get-ADUser $User -Property Title).Title,
    
    [String]$CallBack = "$User | 5-2444 | $UserJob",

    $i=0,
    $j=0
)

    Function SendMessage {

        foreach($Computer in $ComputerName) {

            Write-Progress -Activity "Sending messages..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)         

            #Invoke local MSG command on specified workstation - will generate pop-up message for any user logged onto that workstation - *Also shows on Login screen, stays there for 100,000 seconds or until interacted with
            Invoke-Command -ComputerName $Computer { param($MyMessage, $CallBack, $User, $UserJob)
 
                MSG /time:100000 * /v "$MyMessage {$CallBack}"
            } -ArgumentList $MyMessage, $CallBack, $User, $UserJob -AsJob
        }
    }

    SendMessage | Wait-Job | Remove-Job

}#End NetMSG

Function InstallApplication {

<#     
.SYNOPSIS     
  
    Copies and installs specifed filepath ($Path). This serves as a template for the following filetypes: .EXE, .MSI, & .MSP 

.DESCRIPTION     
    Copies and installs specifed filepath ($Path). This serves as a template for the following filetypes: .EXE, .MSI, & .MSP

.EXAMPLE    
    .\InstallAsJob (Get-Content C:\ComputerList.txt)

.EXAMPLE    
    .\InstallAsJob Computer1, Computer2, Computer3 
    
.NOTES   
    Author: JBear 
    Date: 2/9/2017 
    
    Edit: JBear
    Date: 10/13/2017 
#> 

param(

    [Parameter(Mandatory=$true,HelpMessage="Enter Computername(s)")]
    [String[]]$Computername,

    [Parameter(ValueFromPipeline=$true,HelpMessage="Enter installer path(s)")]
    [String[]]$Path = $null,

    [Parameter(ValueFromPipeline=$true,HelpMessage='Enter remote destination: C$\Directory')]
    $Destination = "C$\TempApplications"
)

    if($Path -eq $null) {

        Add-Type -AssemblyName System.Windows.Forms

        $Dialog = New-Object System.Windows.Forms.OpenFileDialog
        $Dialog.InitialDirectory = "\\lasfs03\Software\Current Version\Deploy"
        $Dialog.Title = "Select Installation File(s)"
        $Dialog.Filter = "Installation Files (*.exe,*.msi,*.msp)|*.exe; *.msi; *.msp"        
        $Dialog.Multiselect=$true
        $Result = $Dialog.ShowDialog()

        if($Result -eq 'OK') {

            Try {
        
                $Path = $Dialog.FileNames
            }

            Catch {

                $Path = $null
	            Break
            }
        }

        else {

            #Shows upon cancellation of Save Menu
            Write-Host -ForegroundColor Yellow "Notice: No file(s) selected."
            Break
        }
    }

    #Create Function    
    Function InstallAsJob {

        #Each item in $Computernam variable        
        foreach($Computer in $Computername) {

            #If $Computer IS NOT null or only whitespace
            if(!([string]::IsNullOrWhiteSpace($Computer))) {

                #Test-Connection to $Computer
                if(Test-Connection -Quiet -Count 1 $Computer) {                                               
                     
                    #Create job on localhost
                    Start-Job { param($Computer, $Path, $Destination)

                        foreach($P in $Path) {
                            
                            #Static Temp location
                            $TempDir = "\\$Computer\$Destination"

                            #Create $TempDir directory
                            if(!(Test-Path $TempDir)) {

                                New-Item -Type Directory $TempDir | Out-Null
                            }
                     
                            #Retrieve Leaf object from $Path
                            $FileName = (Split-Path -Path $P -Leaf)

                            #New Executable Path
                            $Executable = "C:\$(Split-Path -Path $Destination -Leaf)\$FileName"

                            #Copy needed installer files to remote machine
                            Copy-Item -Path $P -Destination $TempDir

                            #Install .EXE
                            if($FileName -like "*.exe") {

                                Function InvokeEXE {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
                                    
                                        Try {

                                            #Start EXE file
                                            Start-Process $Executable -ArgumentList "/s" -Wait -NoNewWindow
                                            
                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }
                                       
                                    } -AsJob -JobName "Silent EXE Install" -ArgumentList $TempDir, $FileName, $Executable
                                }

                                InvokeEXE | Receive-Job -Wait
                            }
                               
                            #Install .MSI                                        
                            elseif($FileName -like "*.msi") {

                                Function InvokeMSI {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
				    
				    	                $MSIArguments = @(
						
						                    "/i"
						                    $Executable
						                    "/qn"
					                    )

                                        Try {
                                        
                                            #Start MSI file                                    
                                            Start-Process 'msiexec.exe' -ArgumentList $MSIArguments -Wait -ErrorAction Stop

                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }                              
                                    } -AsJob -JobName "Silent MSI Install" -ArgumentList $TempDir, $FileName, $Executable                            
                                }

                                InvokeMSI | Receive-Job -Wait
                            }

                            #Install .MSP
                            elseif($FileName -like "*.msp") { 
                                                                       
                                Function InvokeMSP {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
				    
				    	                $MSPArguments = @(
						
						                    "/p"
						                    $Executable
						                    "/qn"
					                    )				    

                                        Try {
                                                                                
                                            #Start MSP file                                    
                                            Start-Process 'msiexec.exe' -ArgumentList $MSPArguments -Wait -ErrorAction Stop

                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }                             
                                    } -AsJob -JobName "Silent MSP Installer" -ArgumentList $TempDir, $FileName, $Executable
                                }

                                InvokeMSP | Receive-Job -Wait
                            }

                            else {

                                Write-Host "$Destination has an unsupported file extension. Please try again."                        
                            }
                        }                      
                    } -Name "Application Install" -Argumentlist $Computer, $Path, $Destination            
                }
                                            
                else {                                
                    
                    Write-Host "Unable to connect to $Computer."                
                }            
            }        
        }   
    }

    #Call main Function
    InstallAsJob
    Write-Host "`nJob creation complete. Please use the Get-Job cmdlet to check progress.`n"
    Write-Host "Once all jobs are complete, use Get-Job | Receive-Job to retrieve any output or, Get-Job | Remove-Job to clear jobs from the session cache."
}#End InstallApplication

# Begin Get-Icon

Function Get-Icon {
    <#
        .SYNOPSIS
            Gets the icon from a file

        .DESCRIPTION
            Gets the icon from a file and displays it in a variety formats.

        .PARAMETER Path
            The path to a file to get the icon

        .PARAMETER ToBytes
            Displays outputs as a byte array

        .PARAMETER ToBitmap
            Display the icon as a bitmap object

        .PARAMETER ToBase64
            Displays the icon in Base64 encoded format

        .NOTES
            Name: Get-Icon
            Author: Boe Prox
            Version History:
                1.0 //Boe Prox - 11JAN2016
                    - Initial version

        .OUTPUT
            System.Drawing.Icon
            System.Drawing.Bitmap
            System.String
            System.Byte[]

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe'

            FullName : C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe
            Handle   : 164169893
            Height   : 32
            Size     : {Width=32, Height=32}
            Width    : 32

            Description
            -----------
            Returns the System.Drawing.Icon representation of the icon

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBitmap

            Tag                  : 
            PhysicalDimension    : {Width=32, Height=32}
            Size                 : {Width=32, Height=32}
            Width                : 32
            Height               : 32
            HorizontalResolution : 96
            VerticalResolution   : 96
            Flags                : 2
            RawFormat            : [ImageFormat: b96b3caa-0728-11d3-9d7b-0000f81ef32e]
            PixelFormat          : Format32bppArgb
            Palette              : System.Drawing.Imaging.ColorPalette
            FrameDimensionsList  : {7462dc86-6180-4c7e-8e3f-ee7333a7a483}
            PropertyIdList       : {}
            PropertyItems        : {}

            Description
            -----------
            Returns the System.Drawing.Bitmap representation of the icon

        .EXAMPLE
            $FileName = 'C:\Temp\PowerShellIcon.png'
            $Format = [System.Drawing.Imaging.ImageFormat]::Png
            (Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBitmap).Save($FileName,$Format)

            Description
            -----------
            Saves the icon as a file.

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBase64

            AAABAAEAICAQHQAAAADoAgAAFgAAACgAAAAgAAAAQAAAAAEABAAAAAAAgAIAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAACAAACAAAAAgIAAgAAAAIAAgACAgAAAgICAAMDAwAAAAP8AAP8AAAD//wD/AAAA/wD/AP
            //AAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABmZmZmZmZmZmZmZgAAAAAAaId3d3d3d4iIiIdgAA
            AHdmhmZmZmZmZmZmZoZAAAB2ZnZmZmZmZmZmZmZ3YAAAdmZ3ZmiHZniIiHZmaGAAAHZmd2Zv/4eIiIi
            GZmhgAAB2ZmdmZ4/4eIh3ZmZnYAAAd2ZnZmZo//h2ZmZmZ3YAAHZmaGZmZo//h2ZmZmd2AAB3Zmd2Zm
            Znj/h2ZmZmhgAAd3dndmZmZuj/+GZmZoYAAHd3dod3dmZuj/9mZmZ2AACHd3aHd3eIiP/4ZmZmd2AAi
            Hd2iIiIiI//iId2ZndgAIiIhoiIiIj//4iIiIiIYACIiId4iIiP//iIiIiIiGAAiIiIaIiI//+IiIiI
            iIhkAIiIiGiIiP/4iIiIiIiIdgCIiIhoiIj/iIiIiIiIiIYAiIiIeIiIiIiIiIiIiIiGAAiIiIaP///
            ////////4hgAAAAAGZmZmZmZmZmZmZmYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD////////////////gA
            AAf4AAAD+AAAAfgAAAHAAAABwAAAAcAAAAHAAAAAwAAAAMAAAADAAAAAwAAAAMAAAABAAAAAQAAAAEA
            AAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAP4AAAH//////////////////////////w==

            Description
            -----------
            Returns the Base64 encoded representation of the icon

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBase64 | Clip

            Description
            -----------
            Returns the Base64 encoded representation of the icon and saves it to the clipboard.

        .EXAMPLE
            (Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBytes) -Join ''

            0010103232162900002322002200040000320006400010400000128200000000000000000000000
            0128001280001281280128000128012801281280012812812801921921920002550025500025525
            5025500025502550255255002552552550000000000000000000000000000000000000000000000
            0000000000000000000000000000000000006102102102102102102102102102102960000613611
            9119119119119120136136136118000119102134102102102102102102102102102134640011810
            2118102102102102102102102102102119960011810211910210413510212013613611810210496
            0011810211910211125513513613613613410210496001181021031021031432481201361191021
            0210396001191021031021021042552481181021021021031180011810210410210210214325513
            5102102102103118001191021031181021021031432481181021021021340011911910311810210
            2102232255248102102102134001191191181351191181021101432551021021021180013511911
            8135119119136136255248102102102119960136119118136136136136143255136135118102119
            9601361361341361361361362552551361361361361369601361361351201361361432552481361
            3613613613696013613613610413613625525513613613613613613610001361361361041361362
            5524813613613613613613611801361361361041361362551361361361361361361361340136136
            1361201361361361361361361361361361361340813613613414325525525525525525525525524
            8134000061021021021021021021021021021021020000000000000000000000000000000000000
            0000000000000000000000000000000000000000000025525525525525525525525525525525525
            5224003122400152240072240070007000700070003000300030003000300010001000100010000
            0000000000000000000012800025400125525525525525525525525525525525525525525525525
            5255255255255

            Description
            -----------
            Returns the bytes representation of the icon. -Join was used in this for the sake
            of displaying all of the data.

    #>
    [cmdletbinding(
        DefaultParameterSetName = '__DefaultParameterSetName'
    )]
    Param (
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullorEmpty()]
        [string]$Path,
        [parameter(ParameterSetName = 'Bytes')]
        [switch]$ToBytes,
        [parameter(ParameterSetName = 'Bitmap')]
        [switch]$ToBitmap,
        [parameter(ParameterSetName = 'Base64')]
        [switch]$ToBase64
    )
    Begin {
        If ($PSBoundParameters.ContainsKey('Debug')) {
            $DebugPreference = 'Continue'
        }
        Add-Type -AssemblyName System.Drawing
    }
    Process {
        $Path = Convert-Path -Path $Path
        Write-Debug $Path
        If (Test-Path -Path $Path) {
            $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path)| 
            Add-Member -MemberType NoteProperty -Name FullName -Value $Path -PassThru
            If ($PSBoundParameters.ContainsKey('ToBytes')) {
                Write-Verbose "Retrieving bytes"
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $MemoryStream.ToArray()   
                $MemoryStream.Flush()  
                $MemoryStream.Dispose()           
            } ElseIf ($PSBoundParameters.ContainsKey('ToBitmap')) {
                $Icon.ToBitMap()
            } ElseIf ($PSBoundParameters.ContainsKey('ToBase64')) {
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $Bytes = $MemoryStream.ToArray()   
                $MemoryStream.Flush() 
                $MemoryStream.Dispose()
                [convert]::ToBase64String($Bytes)
            }  Else {
                $Icon
            }
        } Else {
            Write-Warning "$Path does not exist!"
            Continue
        }
    }
}

# End Get-Icon

# Get Mapped Drive
Function Get-MappedDrive {
	param (
	    [string]$computername = "localhost"
	)
	    Get-WmiObject -Class Win32_MappedLogicalDisk -ComputerName $computername | 
	    Format-List DeviceId, VolumeName, SessionID, Size, FreeSpace, ProviderName
	}
# End Get Mapped Drive

# LayZ LazyWinAdmin GUI Tool
Function LayZ {
    C:\LazyWinAdmin\LazyWinAdmin\LazyWinAdmin.ps1
    }
# End LayZ LazyWinAdmin GUI Tool

# User Last Login
Function Get-UserLastLogonTime{

<#
.SYNOPSIS
Gets the last logon time of users on a Computer.

.DESCRIPTION
Pulls information from the wmi object Win32_UserProfile and outputs an array of objects with properties Name and LastUseTime.
If a date that is year 1 is outputted, then an error occured.

.PARAMETER ComputerName
[object] Specify which computer to target when finding logged on Users.
Default is the host computer

.PARAMETER User
[string] Specify a user to find on the computer.

.PARAMETER ListAllUsers
[switch] Specify the Function to list all users that logged into the computer.

.PARAMETER GetLastUsers
[switch] Specify the Function to get the last user to log onto the computer.

.PARAMETER ListCommonUsers
[switch] Specify to the Function to list common user.

.INPUTS
You may pipe objects into the ComputerName parameter.

.OUTPUTS
outputs an object array with a size dependant on the number of users that logged in with propeties Name and LastUseTime.


#>

    [cmdletBinding()]
    param(
        #computer Name
        [parameter(Mandatory = $False, ValueFromPipeline = $True)]
        [object]$ComputerName = $env:COMPUTERNAME,

        #parameter set, can only choose one from this group
        [parameter(Mandatory = $False, parameterSetName = 'user')]
        [string] $User,
        [parameter(ParameterSetName = 'all users')]
        [switch] $ListAllUsers,
        [parameter(ParameterSetName = 'Last user')]
        [switch] $GetLastUser,

        #Whether or not you want the Function to list Common users
        [switch] $ListCommonUsers
    )

    #Begin Pipeline
    Begin{
        #List of users that are present on all PCs, these won't output unless specified to do so.
        $CommonUsers = ("NetworkService", "LocalService", "systemprofile")
    }
    
    #Process Pipeline
    Process{
        #ping the machine before trying to do anything
        if(Test-Connection $ComputerName -Count 2 -Quiet){
            #try to get the OS version of the computer
            try{$OS = (gwmi  -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction stop).caption}
            catch{
                #had an error getting the WMI-object
                return New-Object psobject -Property @{
                            User = "Error getting WMIObject Win32_OperatingSystem"
                            LastUseTime = get-date 0
                            }
              }
            #make sure the OS retrieved is either Windows 7 or Windows 10 as this Function has not been set to work on other operating systems
            if($OS.contains("Windows 10") -or $OS.Contains("Windows 7")){
                try{
                    #try to get the WMiObject win32_UserProfile
                    $UserObjects = Get-WmiObject -ComputerName $ComputerName -Class Win32_userProfile -Property LocalPath,LastUseTime -ErrorAction Stop
                    $users = @()
                    #loop that handles all the data that came back from the WMI objects
                    forEach($UserObject in $UserObjects){
                        #extract the username from the local path, first find where the last slash is after, get everything after that into its own string
                        $i = $UserObject.localPath.LastIndexOf("\") + 1
                        $tempUserString = ""
                        while($UserObject.localPath.toCharArray()[$i] -ne $null){
                            $tempUserString += $UserObject.localPath.toCharArray()[$i]
                            $i++
                        }
                        #if list common users is turned on, skip this block
                        if(!$listCommonUsers){
                            #check if the user extracted from the local path is a common user
                            $isCommonUser = $false
                            forEach($userName in $CommonUsers){ 
                                if($userName -eq $tempUserString){
                                    $isCommonUser = $true
                                    break
                                }
                            }
                        }
                        #if the user is one of the users specified in the common users, skip it unless otherwise specified
                        if($isCommonUser){continue}
                        #check to see if the user has a timestamp for there last logon 
                        if($UserObject.LastUseTime -ne $null){
                            #This converts the string timestamp to a DateTime object
                            $TempUserLastUseTime = ([WMI] '').ConvertToDateTime($userObject.LastUseTime)
                        }
                        #the user had a local path but no timestamp was found for last logon
                        else{$TempUserLastUseTime = Get-Date 0}
                        #add user to array
                        $users += New-Object -TypeName psobject -Property @{
                            User = $tempUserString
                            LastUseTime = $TempUserLastUseTime
                            }
                    }
                }
                catch{
                    #error trying to retrieve WMI obeject win32_userProfile
                    return New-Object psobject -Property @{
                        User = "Error getting WMIObject Win32_userProfile"
                        LastUseTime = get-date 0
                        }
                }
            }
            else{
                #OS version was not compatible
                return New-Object psobject -Property @{
                    User = "Operating system $OS is not compatible with this Function."
                    LastUseTime = get-date 0
                    }
            }
        }
        else{
            #Computer was not pingable
            return New-Object psobject -Property @{
                User = "Can't Ping"
                LastUseTime = get-date 0
                }
        }

        #check to see if any users came out of the main Function
        if($users.count -eq 0){
            $users += New-Object -TypeName psobject -Property @{
                User = "NoLoggedOnUsers"
                LastUseTime = get-date 0
            }
        }
        #sort the user array by the last time they logged in
        else{$users = $users | Sort-Object -Property LastUseTime -Descending}
        #main output block
        #if List all users was chosen, output the full list of users found
        if($ListAllUsers){return $users}
        #if get last user was chosen, output the last user to log on the computer
        elseif($GetLastUser){return ($users[0])}
        else{
            #see if the user specified ever logged on
            ForEach($Username in $users){
                if($Username.User -eq $user) {return ($Username)}            
            }

            #user did not log on
            return New-Object psobject -Property @{
                User = "$user"
                LastUseTime = get-date 0
                }
        }
    }
    #End Pipeline
    End{Write-Verbose "Function get-UserLastLogonTime is complete"}
}
# End User Last Login

# Begin Unblock
Function Unblock ($path) { 

Get-ChildItem "$path" -Recurse | Unblock-File

}
# End Unblock 

# Begin Get-RemoteSysInfo
Function Get-RemoteSysInfo {
  <# 
  .SYNOPSIS 
  Retrieve basic system information for specified workstation(s) 

  .EXAMPLE 
  Get-RemoteSysInfo Computer123456 

  .EXAMPLE 
  Get-RemoteSysInfo 123456 
  #> 
param(

    [Parameter(Mandatory=$true)]
    [string[]] $ComputerName
)

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

$i=0
$j=0

Function Systeminformation {
	
    foreach ($Computer in $ComputerName) {

        if(!([String]::IsNullOrWhiteSpace($Computer))) {

            If (Test-Connection -quiet -count 1 -Computer $Computer) {

                Write-Progress -Activity "Getting Sytem Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	            Start-Job -ScriptBlock { param($Computer) 

	                #Gather specified workstation information; CimInstance only works on 64-bit
	                $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer
	                $computerBIOS = Get-CimInstance CIM_BIOSElement -Computer $Computer
	                $computerOS = Get-CimInstance CIM_OperatingSystem -Computer $Computer
	                $computerCPU = Get-CimInstance CIM_Processor -Computer $Computer
	                $computerHDD = Get-CimInstance Win32_LogicalDisk -Computer $Computer -Filter "DeviceID = 'C:'"
    
                        [pscustomobject]@{

                            "Computer Name"=$computerSystem.Name
                            "Last Reboot"=$computerOS.LastBootUpTime
                            "Operating System"=$computerOS.OSArchitecture + " " + $computerOS.caption
                             Model=$computerSystem.Model
                             RAM= "{0:N2}" -f [int]($computerSystem.TotalPhysicalMemory/1GB) + "GB"
                            "Disk Capacity"="{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
                            "Total Disk Space"="{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
                            "Current User"=$computerSystem.UserName
                        }
	            } -ArgumentList $Computer
            }

            else {

                Start-Job -ScriptBlock { param($Computer)  
                     
                    [pscustomobject]@{

                        "Computer Name"=$Computer
                        "Last Reboot"="Unable to PING."
                        "Operating System"="$Null"
                        Model="$Null"
                        RAM="$Null"
                        "Disk Capacity"="$Null"
                        "Total Disk Space"="$Null"
                        "Current User"="$Null"
                    }
                } -ArgumentList $Computer                       
            }
        }

        else {
                 
            Start-Job -ScriptBlock { param($Computer)  
                     
                [pscustomobject]@{

                    "Computer Name"="Value is null."
                    "Last Reboot"="$Null"
                    "Operating System"="$Null"
                    Model="$Null"
                    RAM="$Null"
                    "Disk Capacity"="$Null"
                    "Total Disk Space"="$Null"
                    "Current User"="$Null"
                }
            } -ArgumentList $Computer
        }
    } 
}

$SystemInformation = SystemInformation | Wait-Job | Receive-Job | Select "Computer Name", "Current User", "Operating System", Model, RAM, "Disk Capacity", "Total Disk Space", "Last Reboot"
$DocPath = [environment]::getfolderpath("mydocuments") + "\SystemInformation-Report.csv"

	Switch ($CheckBox.IsChecked){
		$true { $SystemInformation | Export-Csv $DocPath -NoTypeInformation -Force; }
		default { $SystemInformation | Out-GridView -Title "System Information"; }
		
    }

	if ($CheckBox.IsChecked -eq $true){

	    Try { 

		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {

		 #Do Nothing 
	    }
	}
	
	else{

	    Try {

	        $listBox.Items.Add("$stamp System Information output processed!`n")
	    } 

	    Catch {

	        #Do Nothing 
	    }
	}
}
# End Get-RemoteSysInfo
#Begin Get-RemoteSoftWare
Function Get-RemoteSoftWare {
  <# 
  .SYNOPSIS 
  Grabs all installed Software on specified computer(s) 

  .EXAMPLE 
  Get-RemoteSoftWare Computer123456 

  .EXAMPLE 
  Get-RemoteSoftWare 123456 
  #> 
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = '')
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

Function SoftwareCheck {

$i=0
$j=0

foreach ($computer in $ComputerArray) {

    Write-Progress -Activity "Retrieving Software Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

        $keys = '','\Wow6432Node'
        foreach ($key in $keys) {
            try {
                $apps = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$computer).OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall").GetSubKeyNames()
            } catch {
                continue
            }

            foreach ($app in $apps) {
                $program = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$computer).OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall\$app")
                $name = $program.GetValue('DisplayName')
                if ($name -and $name -match $NameRegex) {
                    [pscustomobject]@{
                        "Computer Name" = $computer
                        Software = $name
                        Version = $program.GetValue('DisplayVersion')
                        Publisher = $program.GetValue('Publisher')
                        "Install Date" = $program.GetValue('InstallDate')
                        "Uninstall String" = $program.GetValue('UninstallString')
                        Bits = $(if ($key -eq '\Wow6432Node') {'64'} else {'32'})
                        Path = $program.name
                    }
                }
            }
        } 
    }
}	

foreach ($computer in $ComputerName) {	     
    If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
        $ComputerArray += $Computer
    }	
}
	$SoftwareCheck = SoftwareCheck | Sort "Computer Name" | Select "Computer Name", Software, Version, Publisher, "Install Date", "Uninstall String", Bits, Path
    	$DocPath = [environment]::getfolderpath("mydocuments") + "\Software-Report.csv"

    		Switch ($CheckBox.IsChecked){
    		    $true { $SoftwareCheck | Export-Csv $DocPath -NoTypeInformation -Force; }
    		    default { $SoftwareCheck | Out-GridView -Title "Software"; }
		}
		
	if ($CheckBox.IsChecked -eq $true){
	    Try { 
		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {
		 #Do Nothing 
	    }
	}
	
	else{
	    Try {
	        $listBox.Items.Add("$stamp Software output processed!`n")
	    } 
	    Catch {
	        #Do Nothing 
	    }
	}
}
# End Get-RemoteSoftWare
# Begin Get-OfficeVersion
Function Get-OfficeVersion {
<#
.Synopsis
Gets the Office Version installed on the computer
.DESCRIPTION
This Function will query the local or a remote computer and return the information about Office Products installed on the computer
.NOTES   
Name: Get-OfficeVersion
Version: 1.0.5
DateCreated: 2015-07-01
DateUpdated: 2016-07-20
.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
.PARAMETER ComputerName
The computer or list of computers from which to query 
.PARAMETER ShowAllInstalledProducts
Will expand the output to include all installed Office products
.EXAMPLE
Get-OfficeVersion
Description:
Will return the locally installed Office product
.EXAMPLE
Get-OfficeVersion -ComputerName client01,client02
Description:
Will return the installed Office product on the remote computers
.EXAMPLE
Get-OfficeVersion | select *
Description:
Will return the locally installed Office product with all of the available properties
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"

    $excelKeyPath = "Excel\DefaultIcon"
    $wordKeyPath = "Word\DefaultIcon"
   
    $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

    $officeKeys = 'SOFTWARE\Microsoft\Office',
                  'SOFTWARE\Wow6432Node\Microsoft\Office'

    $defaultDisplaySet = 'DisplayName','Version', 'ComputerName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

process {

 $results = new-object PSObject[] 0;

 foreach ($computer in $ComputerName) {
    if ($Credentials) {
       $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
    } else {
       $os=Get-WMIObject win32_operatingsystem -computername $computer
    }

    $osArchitecture = $os.OSArchitecture

    if ($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
    } else {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer
    }

    [System.Collections.ArrayList]$VersionList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PackageList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ConfigItemList = New-Object -TypeName  System.Collections.ArrayList
    $ClickToRunList = new-object PSObject[] 0;

    foreach ($regKey in $officeKeys) {
       $officeVersion = $regProv.EnumKey($HKLM, $regKey)
       foreach ($key in $officeVersion.sNames) {
          if ($key -match "\d{2}\.\d") {
            if (!$VersionList.Contains($key)) {
              $AddItem = $VersionList.Add($key)
            }

            $path = join-path $regKey $key

            $configPath = join-path $path "Common\Config"
            $configItems = $regProv.EnumKey($HKLM, $configPath)
            if ($configItems) {
               foreach ($configId in $configItems.sNames) {
                 if ($configId) {
                    $Add = $ConfigItemList.Add($configId.ToUpper())
                 }
               }
            }

            $cltr = New-Object -TypeName PSObject
            $cltr | Add-Member -MemberType NoteProperty -Name InstallPath -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name UpdatesEnabled -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name UpdateUrl -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name StreamingFinished -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name Platform -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name ClientCulture -Value ""
            
            $packagePath = join-path $path "Common\InstalledPackages"
            $clickToRunPath = join-path $path "ClickToRun\Configuration"
            $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue

            [string]$officeLangResourcePath = join-path  $path "Common\LanguageResources"
            $mainLangId = $regProv.GetDWORDValue($HKLM, $officeLangResourcePath, "SKULanguage").uValue
            if ($mainLangId) {
                $mainlangCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $mainLangId}
                if ($mainlangCulture) {
                    $cltr.ClientCulture = $mainlangCulture.Name
                }
            }

            [string]$officeLangPath = join-path  $path "Common\LanguageResources\InstalledUIs"
            $langValues = $regProv.EnumValues($HKLM, $officeLangPath);
            if ($langValues) {
               foreach ($langValue in $langValues) {
                  $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $langValue}
               } 
            }

            if ($virtualInstallPath) {

            } else {
              $clickToRunPath = join-path $regKey "ClickToRun\Configuration"
              $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue
            }

            if ($virtualInstallPath) {
               if (!$ClickToRunPathList.Contains($virtualInstallPath.ToUpper())) {
                  $AddItem = $ClickToRunPathList.Add($virtualInstallPath.ToUpper())
               }

               $cltr.InstallPath = $virtualInstallPath
               $cltr.StreamingFinished = $regProv.GetStringValue($HKLM, $clickToRunPath, "StreamingFinished").sValue
               $cltr.UpdatesEnabled = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdatesEnabled").sValue
               $cltr.UpdateUrl = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdateUrl").sValue
               $cltr.Platform = $regProv.GetStringValue($HKLM, $clickToRunPath, "Platform").sValue
               $cltr.ClientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, "ClientCulture").sValue
               $ClickToRunList += $cltr
            }

            $packageItems = $regProv.EnumKey($HKLM, $packagePath)
            $officeItems = $regProv.EnumKey($HKLM, $path)

            foreach ($itemKey in $officeItems.sNames) {
              $itemPath = join-path $path $itemKey
              $installRootPath = join-path $itemPath "InstallRoot"

              $filePath = $regProv.GetStringValue($HKLM, $installRootPath, "Path").sValue
              if (!$PathList.Contains($filePath)) {
                  $AddItem = $PathList.Add($filePath)
              }
            }

            foreach ($packageGuid in $packageItems.sNames) {
              $packageItemPath = join-path $packagePath $packageGuid
              $packageName = $regProv.GetStringValue($HKLM, $packageItemPath, "").sValue
            
              if (!$PackageList.Contains($packageName)) {
                if ($packageName) {
                   $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
                }
              }
            }

          }
       }
    }

    

    foreach ($regKey in $installKeys) {
        $keyList = new-object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
           $path = join-path $regKey $key
           $installPath = $regProv.GetStringValue($HKLM, $path, "InstallLocation").sValue
           if (!($installPath)) { continue }
           if ($installPath.Length -eq 0) { continue }

           $buildType = "64-Bit"
           if ($osArchitecture -eq "32-bit") {
              $buildType = "32-Bit"
           }

           if ($regKey.ToUpper().Contains("Wow6432Node".ToUpper())) {
              $buildType = "32-Bit"
           }

           if ($key -match "{.{8}-.{4}-.{4}-1000-0000000FF1CE}") {
              $buildType = "64-Bit" 
           }

           if ($key -match "{.{8}-.{4}-.{4}-0000-0000000FF1CE}") {
              $buildType = "32-Bit" 
           }

           if ($modifyPath) {
               if ($modifyPath.ToLower().Contains("platform=x86")) {
                  $buildType = "32-Bit"
               }

               if ($modifyPath.ToLower().Contains("platform=x64")) {
                  $buildType = "64-Bit"
               }
           }

           $primaryOfficeProduct = $false
           $officeProduct = $false
           foreach ($officeInstallPath in $PathList) {
             if ($officeInstallPath) {
                $installReg = "^" + $installPath.Replace('\', '\\')
                $installReg = $installReg.Replace('(', '\(')
                $installReg = $installReg.Replace(')', '\)')
                if ($officeInstallPath -match $installReg) { $officeProduct = $true }
             }
           }

           if (!$officeProduct) { continue };
           
           $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue          

           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE") -and $name.ToUpper() -notlike "*MUI*" -and $name.ToUpper() -notlike "*VISIO*" -and $name.ToUpper() -notlike "*PROJECT*") {
              $primaryOfficeProduct = $true
           }

           $clickToRunComponent = $regProv.GetDWORDValue($HKLM, $path, "ClickToRunComponent").uValue
           $uninstallString = $regProv.GetStringValue($HKLM, $path, "UninstallString").sValue
           if (!($clickToRunComponent)) {
              if ($uninstallString) {
                 if ($uninstallString.Contains("OfficeClickToRun")) {
                     $clickToRunComponent = $true
                 }
              }
           }

           $modifyPath = $regProv.GetStringValue($HKLM, $path, "ModifyPath").sValue 
           $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue

           $cltrUpdatedEnabled = $NULL
           $cltrUpdateUrl = $NULL
           $clientCulture = $NULL;

           [string]$clickToRun = $false

           if ($clickToRunComponent) {
               $clickToRun = $true
               if ($name.ToUpper().Contains("MICROSOFT OFFICE")) {
                  $primaryOfficeProduct = $true
               }

               foreach ($cltr in $ClickToRunList) {
                 if ($cltr.InstallPath) {
                   if ($cltr.InstallPath.ToUpper() -eq $installPath.ToUpper()) {
                       $cltrUpdatedEnabled = $cltr.UpdatesEnabled
                       $cltrUpdateUrl = $cltr.UpdateUrl
                       if ($cltr.Platform -eq 'x64') {
                           $buildType = "64-Bit" 
                       }
                       if ($cltr.Platform -eq 'x86') {
                           $buildType = "32-Bit" 
                       }
                       $clientCulture = $cltr.ClientCulture
                   }
                 }
               }
           }
           
           if (!$primaryOfficeProduct) {
              if (!$ShowAllInstalledProducts) {
                  continue
              }
           }

           $object = New-Object PSObject -Property @{DisplayName = $name; Version = $version; InstallPath = $installPath; ClickToRun = $clickToRun; 
                     Bitness=$buildType; ComputerName=$computer; ClickToRunUpdatesEnabled=$cltrUpdatedEnabled; ClickToRunUpdateUrl=$cltrUpdateUrl;
                     ClientCulture=$clientCulture }
           $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
           $results += $object

        }
    }

  }

  $results = Get-Unique -InputObject $results 

  return $results;
}

}
# End Get-OfficeVersion
# Begin Get-OfficeVersion2

Function Get-OfficeVersion2
{
param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [string] $Infile,
    
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [string] $outfile
    )
#$outfile = 'C:\temp\office.csv'
#$infile = 'c:\temp\servers.txt'
Begin
    {
    }
 Process
    {
    $office = @()
    $computers = Get-Content $infile
    $i=0
    $count = $computers.count
    foreach($computer in $computers)
     {
     $i++
     Write-Progress -Activity "Querying Computers" -Status "Computer: $i of $count " `
      -PercentComplete ($i/$count*100)
        $info = @{}
        $version = 0
        try{
          $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer) 
          $reg.OpenSubKey('software\Microsoft\Office').GetSubKeyNames() |% {
            if ($_ -match '(\d+)\.') {
              if ([int]$matches[1] -gt $version) {
                $version = $matches[1]
              }
            }    
          }
          if ($version) {
            Write-Debug("$computer : found $version")
            switch($version) {
                "7" {$officename = 'Office 97' }
                "8" {$officename = 'Office 98' }
                "9" {$officename = 'Office 2000' }
                "10" {$officename = 'Office XP' }
                "11" {$officename = 'Office 97' }
                "12" {$officename = 'Office 2003' }
                "13" {$officename = 'Office 2007' }
                "14" {$officename = 'Office 2010' }
                "15" {$officename = 'Office 2013' }
                "16" {$officename = 'Office 2016' }
                default {$officename = 'Unknown Version'}
            }
    
          }
          }
          catch{
              $officename = 'Not Installed/Not Available'
          }
    $info.Computer = $computer
    $info.Name= $officename
    $info.version =  $version

    $object = new-object -TypeName PSObject -Property $info
    $office += $object
    }
    $office | select computer,version,name | Export-Csv -NoTypeInformation -Path $outfile
    }
}
  write-output ("Done")
  # End Get-OfficeVersion2
  # Begin Get-OutlookClientVersion
  
Function Get-OutlookClientVersion {
 
<#
.SYNOPSIS
    Identifies and reports which Outlook client versions are being used to access Exchange.
 
.DESCRIPTION
    Get-MrRCAProtocolLog is an advanced PowerShell Function that parses Exchange Server RPC
    logs to determine what Outlook client versions are being used to access the Exchange Server.
 
.PARAMETER LogFile
    The path to the Exchange RPC log files.
 
.EXAMPLE
     Get-MrRCAProtocolLog -LogFile 'C:\Program Files\Microsoft\Exchange Server\V15\Logging\RPC Client Access\RCA_20140831-1.LOG'
 
.EXAMPLE
     Get-ChildItem -Path '\\servername\c$\Program Files\Microsoft\Exchange Server\V15\Logging\RPC Client Access\*.log' |
     Get-MrRCAProtocolLog |
     Out-GridView -Title 'Outlook Client Versions'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author:  Mike F Robbins
    Website: http://mikefrobbins.com
    Twitter: @mikefrobbins
#>
 
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
                   ValueFromPipeline)]
        [ValidateScript({
            Test-Path -Path $_ -PathType Leaf -Include '*.log'
        })]
        [string[]]$LogFile
    )
 
    PROCESS {
        foreach ($file in $LogFile) {
            $Headers = (Get-Content -Path $file -TotalCount 5 | Where-Object {$_ -like '#Fields*'}) -replace '#Fields: ' -split ','
                    
            Import-Csv -Header $Headers -Path $file |
            Where-Object {$_.operation -eq 'Connect' -and $_.'client-software' -eq 'outlook.exe'} |
            Select-Object -Unique -Property @{label='User';expression={$_.'client-name' -replace '^.*cn='}},
                                            @{label='DN';expression={$_.'client-name'}},
                                            client-software,
                                            @{label='Version';expression={Get-MrOutlookVersion -OutlookBuild $_.'client-software-version'}},
                                            client-mode,
                                            client-ip,
                                            protocol
        }
    }
}
 
Function Get-MrOutlookVersion {
    param (
        [string]$OutlookBuild
    )
    switch ($OutlookBuild) {  
        {$_ -ge '16.0.4266.1001'} {'Outlook 2016 4266.1001'; break}
        {$_ -ge '16.0.4522.1000'} {'Outlook 2016 4522.1000'; break}
        {$_ -ge '16.0.4498.1000'} {'Outlook 2016 4498.1000'; break}
        {$_ -ge '16.0.4229.1024'} {'Outlook 2016 4229.1024'; break}              
        {$_ -ge '15.0.4569.1506'} {'Outlook 2013 SP1'; break}
        {$_ -ge '15.0.4420.1017'} {'Outlook 2013 RTM'; break}
        {$_ -ge '14.0.7015.1000'} {'Outlook 2010 SP2'; break}
        {$_ -ge '14.0.6029.1000'} {'Outlook 2010 SP1'; break}
        {$_ -ge '14.0.4763.1000'} {'Outlook 2010 RTM'; break}
        {$_ -ge '12.0.6672.5000'} {'Outlook 2007 SP3 U2013'; break}
        {$_ -ge '12.0.6423.1000'} {'Outlook 2007 SP2'; break}
        {$_ -ge '12.0.6212.1000'} {'Outlook 2007 SP1'; break}
        {$_ -ge '12.0.4518.1014'} {'Outlook 2007 RTM'; break}
        Default {'$OutlookBuild'}
    }
}
# End Get-OutlookClientVersion
# Begin VMware Functions

<# Enable or Disable Hot Add Memory/CPU
 Enable-MemHotAdd $ServerName
 Disable-MemHotAdd $ServerName
 Enable-vCPUHotAdd $ServerName
 Disable-vCPUHotAdd $ServerName
#>

Function Enable-MemHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-MemHotAdd($vm){
$vmview = Get-VM $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Enable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End VMware Functions
# Begin Get-DefragAnalysis
Function Get-DefragAnalysis {

<#
.Synopsis
Run a defrag analysis.
.Description
This command uses WMI to to run a defrag analysis on selected volumes on
local or remote computers. You will get a custom object for each volume like
this:

AverageFileSize               : 64
AverageFragmentsPerFile       : 1
AverageFreeSpacePerExtent     : 17002496
ClusterSize                   : 4096
ExcessFolderFragments         : 0
FilePercentFragmentation      : 0
FragmentedFolders             : 0
FreeSpace                     : 161816576
FreeSpacePercent              : 77
FreeSpacePercentFragmentation : 29
LargestFreeSpaceExtent        : 113500160
MFTPercentInUse               : 100
MFTRecordCount                : 511
PageFileSize                  : 0
TotalExcessFragments          : 0
TotalFiles                    : 182
TotalFolders                  : 11
TotalFragmentedFiles          : 0
TotalFreeSpaceExtents         : 8
TotalMFTFragments             : 1
TotalMFTSize                  : 524288
TotalPageFileFragments        : 0
TotalPercentFragmentation     : 0
TotalUnmovableFiles           : 4
UsedSpace                     : 47894528
VolumeName                    : 
VolumeSize                    : 209711104
Driveletter                   : E:
DefragRecommended             : False
Computername                  : NOVO8

The default drive is C: on the local computer.
.Example
PS C:\> Get-DefragAnalysis
Run a defrag analysis on C: on the local computer
.Example
PS C:\> Get-DefragAnalysis -drive "C:" -computername $servers
Run a defrag analysis for drive C: on a previously defined collection of server names.
.Example
PS C:\> $data = Get-WmiObject Win32_volume -filter "driveletter like '%' AND drivetype=3" -ComputerName Novo8 | Get-DefragAnalysis
PS C:\> $data | Sort Driveletter | Select Computername,DriveLetter,DefragRecommended

Computername                    Driveletter                     DefragRecommended
------------                    -----------                     -----------------
NOVO8                           C:                                          False
NOVO8                           D:                                          True
NOVO8                           E:                                          False

Get all volumes on a remote computer that are fixed but have a drive letter,
this should eliminate CD/DVD drives, and run a defrag analysis on each one.
The results are saved to a variable, $data.
.Notes
Last Updated: 12/5/2012
Author      : Jeffery Hicks (http://jdhitsolutions.com/blog)
Version     : 0.9

.Link
Get-WMIObject
Invoke-WMIMethod

#>

[cmdletbinding(SupportsShouldProcess=$True)]

Param(
[Parameter(Position=0,ValueFromPipelineByPropertyName=$True)]
[ValidateNotNullorEmpty()]
[Alias("drive")]
[string]$Driveletter="C:",
[Parameter(Position=1,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[ValidateNotNullorEmpty()]
[Alias("PSComputername","SystemName")]
[string[]]$Computername=$env:computername
)

Begin {
    Write-Verbose -Message "$(Get-Date) Starting $($MyInvocation.Mycommand)"   
} #close Begin

Process {
    #strip off any extra spaces on the drive letter just in case
    Write-Verbose "$(Get-Date) Processing $Driveletter"
    $Driveletter=$Driveletter.Trim()
    if ($Driveletter.length -gt 2) {
        Write-Verbose "$(Get-Date) Scrubbing drive parameter value"
        $Driveletter=$Driveletter.Substring(0,2)
    }
    #add a colon if not included
    if ($Driveletter -match "^\w$") {
        Write-Verbose "$(Get-Date) Modifying drive parameter value"
        $Driveletter="$($Driveletter):"
    }

    Write-Verbose "$(Get-Date) Analyzing drive $Driveletter"
        
    Foreach ($computer in $computername) {
        Write-Verbose "$(Get-Date) Examining $computer"
        Try {
            $volume=Get-WmiObject -Class Win32_Volume -filter "DriveLetter='$Driveletter'" -computername $computer -errorAction "Stop"
        }
        Catch {
            Write-Warning ("Failed to get volume {0} from  {1}. {2}" -f $driveletter,$computer,$_.Exception.Message)
        }
        if ($volume) {
            Write-Verbose "$(Get-Date) Running defrag analysis"
            $analysis = $volume | Invoke-WMIMethod -name DefragAnalysis
        
            #get properties for DefragAnalysis so we can filter out system properties
            $analysis.DefragAnalysis.Properties | 
            Foreach -begin {$Prop=@()} -process { $Prop+=$_.Name }
        
            Write-Verbose "$(Get-Date) Retrieving results"
            $analysis | Select @{Name="Results";Expression={$_.DefragAnalysis | 
            Select-Object -Property $Prop |
            Foreach-Object { 
              #Add on some additional property values
              $_ | Add-member -MemberType Noteproperty -Name Driveletter -value $DriveLetter
              $_ | Add-member -MemberType Noteproperty -Name DefragRecommended -value $analysis.DefragRecommended 
              $_ | Add-member -MemberType Noteproperty -Name Computername -value $volume.__SERVER -passthru
             } #foreach-object
            }}  | Select -expand Results 
            
            #clean up variables so there are no accidental leftovers
            Remove-Variable "volume","analysis"
        } #close if volume
     } #close Foreach computer
 } #close Process
 
End {
    Write-Verbose "$(Get-Date) Defrag analysis complete"
} #close End

} #close Function
# End Get-DefragAnalysis
# Begin Get-NetworkInfo
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
# End Get-NetworkInfo
# Begin Get-SNMPTrap
Function Get-SnmpTrap {
<#
.SYNOPSIS
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
# End Get-SNMPTrap
# Begin Get-UserLogon
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
 
$ErrorActionPreference="SilentlyContinue"
 
$result=@()
 
If ($Computer) {
 
Invoke-Command -ComputerName $Computer -ScriptBlock {quser} | Select-Object -Skip 1 | Foreach-Object {
 
$b=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($b[2] -like 'Disc*') {
 
$array= ([ordered]@{
'User' = $b[0]
'Computer' = $Computer
'Date' = $b[4]
'Time' = $b[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
else {
 
$array= ([ordered]@{
'User' = $b[0]
'Computer' = $Computer
'Date' = $b[5]
'Time' = $b[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
}
}
 
If ($OU) {
 
$comp=Get-ADComputer -Filter * -SearchBase "$OU" -Properties operatingsystem
 
$count=$comp.count
 
If ($count -gt 20) {
 
Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer"
 
}
 
foreach ($u in $comp) {
 
Invoke-Command -ComputerName $u.Name -ScriptBlock {quser} | Select-Object -Skip 1 | ForEach-Object {
 
$a=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($a[2] -like '*Disc*') {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[4]
'Time' = $a[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
}
 
else {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[5]
'Time' = $a[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
}
 
}
 
}
 
}
 
If ($All) {
 
$comp=Get-ADComputer -Filter * -Properties operatingsystem
 
$count=$comp.count
 
If ($count -gt 20) {
 
Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer ..."
 
}
 
foreach ($u in $comp) {
 
Invoke-Command -ComputerName $u.Name -ScriptBlock {quser} | Select-Object -Skip 1 | ForEach-Object {
 
$a=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($a[2] -like '*Disc*') {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[4]
'Time' = $a[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
else {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[5]
'Time' = $a[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
}
 
}
}
Write-Output $result
}
# End Get-UserLogon
# Begin Invoke-Ping
Function Invoke-Ping
{
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
# Begin New-ISOFile
Function New-IsoFile  
{  
  <#  
   .Synopsis  
    Creates a new .iso file  
   .Description  
    The New-IsoFile cmdlet creates a new .iso file containing content from chosen folders  
   .Example  
    New-IsoFile "c:\tools","c:Downloads\utils"  
    This command creates a .iso file in $env:temp folder (default location) that contains c:\tools and c:\downloads\utils folders. The folders themselves are included at the root of the .iso image.  
   .Example 
    New-IsoFile -FromClipboard -Verbose 
    Before running this command, select and copy (Ctrl-C) files/folders in Explorer first.  
   .Example  
    dir c:\WinPE | New-IsoFile -Path c:\temp\WinPE.iso -BootFile "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\efisys.bin" -Media DVDPLUSR -Title "WinPE" 
    This command creates a bootable .iso file containing the content from c:\WinPE folder, but the folder itself isn't included. Boot file etfsboot.com can be found in Windows ADK. Refer to IMAPI_MEDIA_PHYSICAL_TYPE enumeration for possible media types: http://msdn.microsoft.com/en-us/library/windows/desktop/aa366217(v=vs.85).aspx  
   .Notes 
    NAME:  New-IsoFile  
    AUTHOR: Chris Wu 
    LASTEDIT: 03/23/2016 14:46:50  
 #>  
  
  [CmdletBinding(DefaultParameterSetName='Source')]Param( 
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true, ParameterSetName='Source')]$Source,  
    [parameter(Position=2)][string]$Path = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",  
    [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})][string]$BootFile = $null, 
    [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER', 
    [string]$Title = (Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
    [switch]$Force, 
    [parameter(ParameterSetName='Clipboard')][switch]$FromClipboard 
  ) 
 
  Begin {  
    ($cp = new-object System.CodeDom.Compiler.CompilerParameters).CompilerOptions = '/unsafe' 
    if (!('ISOFile' -as [type])) {  
      Add-Type -CompilerParameters $cp -TypeDefinition @' 
public class ISOFile  
{ 
  public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks)  
  {  
    int bytes = 0;  
    byte[] buf = new byte[BlockSize];  
    var ptr = (System.IntPtr)(&bytes);  
    var o = System.IO.File.OpenWrite(Path);  
    var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;  
  
    if (o != null) { 
      while (TotalBlocks-- > 0) {  
        i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);  
      }  
      o.Flush(); o.Close();  
    } 
  } 
}  
'@  
    } 
  
    if ($BootFile) { 
      if('BDR','BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
      ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type=1}).Open()  # adFileTypeBinary 
      $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
      ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
    } 
 
    $MediaType = @('UNKNOWN','CDROM','CDR','CDRW','DVDROM','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','HDDVDROM','HDDVDR','HDDVDRAM','BDROM','BDR','BDRE') 
 
    Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))" 
    ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName=$Title}).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 
  
    if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break } 
  }  
 
  Process { 
    if($FromClipboard) { 
      if($PSVersionTable.PSVersion.Major -lt 5) { Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break } 
      $Source = Get-Clipboard -Format FileDropList 
    } 
 
    foreach($item in $Source) { 
      if($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) { 
        $item = Get-Item -LiteralPath $item 
      } 
 
      if($item) { 
        Write-Verbose -Message "Adding item to the target image: $($item.FullName)" 
        try { $Image.Root.AddTree($item.FullName, $true) } catch { Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') } 
      } 
    } 
  } 
 
  End {  
    if ($Boot) { $Image.BootImageOptions=$Boot }  
    $Result = $Image.CreateResultImage()  
    [ISOFile]::Create($Target.FullName,$Result.ImageStream,$Result.BlockSize,$Result.TotalBlocks) 
    Write-Verbose -Message "Target image ($($Target.FullName)) has been created" 
    $Target 
  } 
} 
# End New-IsoFile
# Begin Set-FileTime
Function Set-FileTime{
  param(
    [string[]]$paths,
    [bool]$only_modification = $false,
    [bool]$only_access = $false
  );

  begin {
    Function updateFileSystemInfo([System.IO.FileSystemInfo]$fsInfo) {
      $datetime = get-date
      if ( $only_access )
      {
         $fsInfo.LastAccessTime = $datetime
      }
      elseif ( $only_modification )
      {
         $fsInfo.LastWriteTime = $datetime
      }
      else
      {
         $fsInfo.CreationTime = $datetime
         $fsInfo.LastWriteTime = $datetime
         $fsInfo.LastAccessTime = $datetime
       }
    }
   
    Function touchExistingFile($arg) {
      if ($arg -is [System.IO.FileSystemInfo]) {
        updateFileSystemInfo($arg)
      }
      else {
        $resolvedPaths = resolve-path $arg
        foreach ($rpath in $resolvedPaths) {
          if (test-path -type Container $rpath) {
            $fsInfo = new-object System.IO.DirectoryInfo($rpath)
          }
          else {
            $fsInfo = new-object System.IO.FileInfo($rpath)
          }
          updateFileSystemInfo($fsInfo)
        }
      }
    }
   
    Function touchNewFile([string]$path) {
      #$null > $path
      Set-Content -Path $path -value $null;
    }
  }
 
  process {
    if ($_) {
      if (test-path $_) {
        touchExistingFile($_)
      }
      else {
        touchNewFile($_)
      }
    }
  }
 
  end {
    if ($paths) {
      foreach ($path in $paths) {
        if (test-path $path) {
          touchExistingFile($path)
        }
        else {
          touchNewFile($path)
        }
      }
    }
  }
}
# End Set-FileTime
# Begin Get-PendingUpdate
Function Get-PendingUpdate { 
    <#    
      .SYNOPSIS   
        Retrieves the updates waiting to be installed from WSUS   
      .DESCRIPTION   
        Retrieves the updates waiting to be installed from WSUS  
      .PARAMETER Computername 
        Computer or computers to find updates for.   
      .EXAMPLE   
       Get-PendingUpdates 
    
       Description 
       ----------- 
       Retrieves the updates that are available to install on the local system 
      .NOTES 
      Author: Boe Prox                                           
                                        
    #> 
      
    #Requires -version 3.0   
    [CmdletBinding( 
        DefaultParameterSetName = 'computer' 
        )] 
    param( 
        [Parameter(ValueFromPipeline = $True)] 
            [string[]]$Computername = $env:COMPUTERNAME
        )     
    Process { 
        ForEach ($computer in $Computername) { 
            If (Test-Connection -ComputerName $computer -Count 1 -Quiet) { 
                Try { 
                #Create Session COM object 
                    Write-Verbose "Creating COM object for WSUS Session" 
                    $updatesession =  [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$computer)) 
                    } 
                Catch { 
                    Write-Warning "$($Error[0])" 
                    Break 
                    } 
 
                #Configure Session COM Object 
                Write-Verbose "Creating COM object for WSUS update Search" 
                $updatesearcher = $updatesession.CreateUpdateSearcher() 
 
                #Configure Searcher object to look for Updates awaiting installation 
                Write-Verbose "Searching for WSUS updates on client" 
                $searchresult = $updatesearcher.Search("IsInstalled=0")     
             
                #Verify if Updates need installed 
                Write-Verbose "Verifing that updates are available to install" 
                If ($searchresult.Updates.Count -gt 0) { 
                    #Updates are waiting to be installed 
                    Write-Verbose "Found $($searchresult.Updates.Count) update\s!" 
                    #Cache the count to make the For loop run faster 
                    $count = $searchresult.Updates.Count 
                 
                    #Begin iterating through Updates available for installation 
                    Write-Verbose "Iterating through list of updates" 
                    For ($i=0; $i -lt $Count; $i++) { 
                        #Create object holding update 
                        $Update = $searchresult.Updates.Item($i)
                        [pscustomobject]@{
                            Computername = $Computer
                            Title = $Update.Title
                            KB = $($Update.KBArticleIDs)
                            SecurityBulletin = $($Update.SecurityBulletinIDs)
                            MsrcSeverity = $Update.MsrcSeverity
                            IsDownloaded = $Update.IsDownloaded
                            Url = $($Update.MoreInfoUrls)
                            Categories = ($Update.Categories | Select-Object -ExpandProperty Name)
                            BundledUpdates = @($Update.BundledUpdates)|ForEach{
                               [pscustomobject]@{
                                    Title = $_.Title
                                    DownloadUrl = @($_.DownloadContents).DownloadUrl
                                }
                            }
                        } 
                    }
                } 
                Else { 
                    #Nothing to install at this time 
                    Write-Verbose "No updates to install." 
                }
            } 
            Else { 
                #Nothing to install at this time 
                Write-Warning "$($c): Offline" 
            }  
        }
    }  
}
# End Get-PendingUpdate
# Begin Get-Set-NetworkLevelAuthentication
Function Get-NetworkLevelAuthentication
{
<#
	.SYNOPSIS
		This Function will get the NLA setting on a local machine or remote machine

	.DESCRIPTION
		This Function will get the NLA setting on a local machine or remote machine

	.PARAMETER  ComputerName
		Specify one or more computer to query

	.PARAMETER  Credential
		Specify the alternative credential to use. By default it will use the current one.
	
	.EXAMPLE
		Get-NetworkLevelAuthentication
		
		This will get the NLA setting on the localhost
	
		ComputerName     : XAVIERDESKTOP
		NLAEnabled       : True
		TerminalName     : RDP-Tcp
		TerminalProtocol : Microsoft RDP 8.0
		Transport        : tcp	

    .EXAMPLE
		Get-NetworkLevelAuthentication -ComputerName DC01
		
		This will get the NLA setting on the server DC01
	
		ComputerName     : DC01
		NLAEnabled       : True
		TerminalName     : RDP-Tcp
		TerminalProtocol : Microsoft RDP 8.0
		Transport        : tcp
	
	.EXAMPLE
		Get-NetworkLevelAuthentication -ComputerName DC01, SERVER01 -verbose
	
	.EXAMPLE
		Get-Content .\Computers.txt | Get-NetworkLevelAuthentication -verbose
		
	.NOTES
		DATE	: 2014/04/01
		AUTHOR	: Francois-Xavier Cat
		WWW		: http://lazywinadmin.com
		Twitter	: @lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the Function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				[pscustomobject][ordered]@{
					'ComputerName' = $NLAinfo.PSComputerName
					'NLAEnabled' = $NLAinfo.UserAuthenticationRequired -as [bool]
					'TerminalName' = $NLAinfo.TerminalName
					'TerminalProtocol' = $NLAinfo.TerminalProtocol
					'Transport' = $NLAinfo.transport
				}
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{
		
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}


Function Set-NetworkLevelAuthentication
{
<#
	.SYNOPSIS
		This Function will set the NLA setting on a local machine or remote machine

	.DESCRIPTION
		This Function will set the NLA setting on a local machine or remote machine

	.PARAMETER  ComputerName
		Specify one or more computers
	
	.PARAMETER EnableNLA
		Specify if the NetworkLevelAuthentication need to be set to $true or $false
	
	.PARAMETER  Credential
		Specify the alternative credential to use. By default it will use the current one.

	.EXAMPLE
		Set-NetworkLevelAuthentication -EnableNLA $true

		ReturnValue                             PSComputerName                         
		-----------                             --------------                         
		                                        XAVIERDESKTOP      
	
	.NOTES
		DATE	: 2014/04/01
		AUTHOR	: Francois-Xavier Cat
		WWW		: http://lazywinadmin.com
		Twitter	: @lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Parameter(Mandatory)]
		[Bool]$EnableNLA,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
				
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the Function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				$NLAinfo | Invoke-CimMethod -MethodName SetUserAuthenticationRequired -Arguments @{ UserAuthenticationRequired = $EnableNLA } -ErrorAction 'Continue' -ErrorVariable ErrorProcessInvokeWmiMethod
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
				if ($ErrorProcessInvokeWmiMethod) { Write-Warning -Message "PROCESS Error - $ErrorProcessInvokeWmiMethod" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{	
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}
# End Get-Set-NetworkLevelAuthentication
# Begin Get-FolderSize
Function Get-FolderSize 
{

<#
.SYNOPSIS
	Get-FolderSize Function displays size of all folders in a specified path. 
	
.DESCRIPTION

	Get-FolderSize Function allows you to return folders greater than a specified size.
	See examples below for more info. 
	
.PARAMETER FolderPath
	Specifies the path you wish to check folder sizes. For example \\70411SRV\EventLogs
	Will return sizes (in GB) of all folders in \\70411SRV\EventLogs. FolderPath accepts
	both UNC and local path format. You can specify multiple paths in quotes, seperated
	by commas. 
	
.PARAMETER FoldersOver
	This parameter is specified in whole numbers (but represents values in GB). It instructs
	the Get-FolderSize Function to return only folders greater than or equal to the specified
	value in GB. 
	
.PARAMETER Recurse
	If this parameter is specified, size of all folders and subfolders are displayed 
	If the Recurse parameter is not spefified (default), size of base folders are displayed.
	
.EXAMPLE
	To return size for all folders in C:\EventLogs, run the command:
	PS C:\>Get-FolderSize -FolderPath C:\EventLogs
	The command returns the following output:
	Permorning initial tasks, please wait...
	Calculating size of folders in C:\EventLogs. This may take sometime, please wait...
	
	Folder Name                             Full Path                               Size
	-----------                             ---------                               ----
	70411SRV                                C:\EventLogs\70411SRV                   384 MB
	70411SRV1                               C:\EventLogs\70411SRV1                  128 MB
	70411SRV2                               C:\EventLogs\70411SRV2                  128 MB
	70411SRV3                               C:\EventLogs\70411SRV3                  128 MB
	Softwares                               C:\EventLogs\Softwares                  2.34 GB

.EXAMPLE
	To return size for folders in C:\EventLogs greater or equal to 200BM, run the command:
	PS C:\> Get-FolderSize C:\EventLogs -FoldersOver 0.2
	Result of the above command is shown below:
	Permorning initial tasks, please wait...
	Calculating size of folders in C:\EventLogs. This may take sometime, please wait...

	Folder Name                             Full Path                               Size
	-----------                             ---------                               ----
	70411SRV                                C:\EventLogs\70411SRV                   384 MB
	Softwares                               C:\EventLogs\Softwares                  2.34 GB
	Notice that only folders greater than 200 MB were returned 
	
.EXAMPLE
	To return size of all folders and subfolders, specify the Recurse parameter:
	PS C:\> Get-FolderSize C:\EventLogs -Recurse
	
	Performing initial tasks, please wait...
	Calculating size of folders in C:\EventLogs. This may take sometime, please wait...

	Folder Name                             Full Path                               Size
	-----------                             ---------                               ----
	70411SRV                                C:\EventLogs\70411SRV                   384 MB
	70411SRV1                               C:\EventLogs\70411SRV1                  128 MB
	70411SRV2                               C:\EventLogs\70411SRV2                  128 MB
	70411SRV3                               C:\EventLogs\70411SRV3                  128 MB
	Softwares                               C:\EventLogs\Softwares                  2.34 GB
	Acrobat 7                               C:\EventLogs\Softwares\Acrobat 7        209 MB
	Citrix                                  C:\EventLogs\Softwares\Citrix           96 MB
	Dell OM station                         C:\EventLogs\Softwares\Dell OM station  577 MB
	JAWS                                    C:\EventLogs\Softwares\JAWS             227 MB
	MDT 2012 Update1                        C:\EventLogs\Softwares\MDT 2012 Update1 118 MB
	OpenManageEssentials                    C:\EventLogs\Softwares\OpenManageEss... 891 MB
	Adobe Acrobat 7.0 Professional          C:\EventLogs\Softwares\Acrobat 7\Ado... 200 MB
	windows                                 C:\EventLogs\Softwares\Dell OM stati... 271 MB
	ManagementStation                       C:\EventLogs\Softwares\Dell OM stati... 254 MB
	support                                 C:\EventLogs\Softwares\Dell OM stati... 107 MB
	
	Notice that, we now have size of all folders and subfolders
	
	
#>

[CmdletBinding(DefaultParameterSetName='FolderPath')]
param 
(
[Parameter(Mandatory=$true,Position=0,ParameterSetName='FolderPath')]
[String[]]$FolderPath,
[Parameter(Mandatory=$false,Position=1,ParameterSetName='FolderPath')]
[String]$FoldersOver,
[Parameter(Mandatory=$false,Position=2,ParameterSetName='FolderPath')]
[switch]$Recurse

)

Begin 
{
#$FoldersOver and $ZeroSizeFolders cannot be used together
#Convert the size specified by Greaterhan parameter to Bytes
$size = 1000000000 * $FoldersOver

}

Process {#Check whether user has access to the folders.
	
	
		Try {
		Write-Host "Performing initial tasks, please wait... " -ForegroundColor Magenta
		$ColItems = If ($Recurse) {Get-ChildItem $FolderPath -Recurse -ErrorAction Stop } 
		Else {Get-ChildItem $FolderPath -ErrorAction Stop } 
		
		} 
		Catch [exception]{}
		
		#Calculate folder size
		If ($ColItems) 
		{
		Write-Host "Calculating size of folders in $FolderPath. This may take sometime, please wait... " -ForegroundColor Magenta
		$Items = $ColItems | Where-Object {$_.PSIsContainer -eq $TRUE -and `
		@(Get-ChildItem -LiteralPath $_.Fullname -Recurse -ErrorAction SilentlyContinue | Where-Object {!$_.PSIsContainer}).Length -gt '0'}}
		

		ForEach ($i in $Items)
		{

		$subFolders = 
		If ($FoldersOver)
		{Get-ChildItem -Path $i.FullName -Recurse | Measure-Object -sum Length | Where-Object {$_.Sum -ge $size -and $_.Sum -gt 100000000  } }
		Else
		{Get-ChildItem -Path $i.FullName -Recurse | Measure-Object -sum Length | Where-Object {$_.Sum -gt 100000000  }} #added 25/12/2014: returns folders over 100MB
		#Return only values not equal to 0
		ForEach ($subFolder in $subFolders) {
		#If folder is less than or equal to 1GB, display in MB, If above 1GB, display in GB 
		$si = If (($subFolder.Sum -ge 1000000000)  ) {"{0:N2}" -f ($subFolder.Sum / 1GB) + " GB"} 
 	  	ElseIf (($subFolder.Sum -lt 1000000000)  ) {"{0:N0}" -f ($subFolder.Sum / 1MB) + " MB"} 
		$Object = New-Object PSObject -Property @{            
        'Folder Name'    = $i.Name                
        'Size'    =  $si
        'Full Path'    = $i.FullName          
        } 

		$Object | Select-Object 'Folder Name', 'Full Path',Size



} 

}


}
End {

Write-Host "Task completed...if nothing is displayed:
you may not have access to the path specified or 
all folders are less than 100 MB" -ForegroundColor Cyan


}

}
# End Get-FolderSize

# Begin Get-Software
Function Get-Software  {

  [OutputType('System.Software.Inventory')]

  [Cmdletbinding()] 

  Param( 

  [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 

  [String[]]$Computername=$env:COMPUTERNAME

  )         

  Begin {

  }

  Process  {     

  ForEach  ($Computer in  $Computername){ 

  If  (Test-Connection -ComputerName  $Computer -Count  1 -Quiet) {

  $Paths  = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         

  ForEach($Path in $Paths) { 

  Write-Verbose  "Checking Path: $Path"

  #  Create an instance of the Registry Object and open the HKLM base key 

  Try  { 

  $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer,'Registry64') 

  } Catch  { 

  Write-Error $_ 

  Continue 

  } 

  #  Drill down into the Uninstall key using the OpenSubKey Method 

  Try  {

  $regkey=$reg.OpenSubKey($Path)  

  # Retrieve an array of string that contain all the subkey names 

  $subkeys=$regkey.GetSubKeyNames()      

  # Open each Subkey and use GetValue Method to return the required  values for each 

  ForEach ($key in $subkeys){   

  Write-Verbose "Key: $Key"

  $thisKey=$Path+"\\"+$key 

  Try {  

  $thisSubKey=$reg.OpenSubKey($thisKey)   

  # Prevent Objects with empty DisplayName 

  $DisplayName =  $thisSubKey.getValue("DisplayName")

  If ($DisplayName  -AND $DisplayName  -notmatch '^Update  for|rollup|^Security Update|^Service Pack|^HotFix') {

  $Date = $thisSubKey.GetValue('InstallDate')

  If ($Date) {

  Try {

  $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)

  } Catch{

  Write-Warning "$($Computer): $_ <$($Date)>"

  $Date = $Null

  }

  } 

  # Create New Object with empty Properties 

  $Publisher =  Try {

  $thisSubKey.GetValue('Publisher').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('Publisher')

  }

  $Version = Try {

  #Some weirdness with trailing [char]0 on some strings

  $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))

  } 

  Catch {

  $thisSubKey.GetValue('DisplayVersion')

  }

  $UninstallString =  Try {

  $thisSubKey.GetValue('UninstallString').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('UninstallString')

  }

  $InstallLocation =  Try {

  $thisSubKey.GetValue('InstallLocation').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('InstallLocation')

  }

  $InstallSource =  Try {

  $thisSubKey.GetValue('InstallSource').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('InstallSource')

  }

  $HelpLink = Try {

  $thisSubKey.GetValue('HelpLink').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('HelpLink')

  }

  $Object = [pscustomobject]@{

  Computername = $Computer

  DisplayName = $DisplayName

  Version  = $Version

  InstallDate = $Date

  Publisher = $Publisher

  UninstallString = $UninstallString

  InstallLocation = $InstallLocation

  InstallSource  = $InstallSource

  HelpLink = $thisSubKey.GetValue('HelpLink')

  EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))

  }

  $Object.pstypenames.insert(0,'System.Software.Inventory')

  Write-Output $Object

  }

  } Catch {

  Write-Warning "$Key : $_"

  }   

  }

  } Catch  {}   

  $reg.Close() 

  }                  

  } Else  {

  Write-Error  "$($Computer): unable to reach remote system!"

  }

  } 

  } 

}  
# End Get-Software
# Begin Get-AssetTagAndSerialNumber
Function Get-AssetTagAndSerialNumber {

   param  ( [string[]]$computerName = @('.') );

   $computerName | % {

       if ($_) {

           Get-WmiObject -ComputerName $_ Win32_SystemEnclosure | Select-Object __Server, SerialNumber, SMBiosAssetTag

       }

   }

}
# End Get-AssetTagAndSerialNumber
# Begin Enable-MemHotAdd
Function Enable-MemHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End Enable-MemHotAdd
# Begin Disable-MemHotAdd
Function Disable-MemHotAdd($vm){
$vmview = Get-VM $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End Disable-MemHotAdd
# Begin Enable-vCPUHotAdd
Function Enable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End Enable-vCPUHotAdd
# Begin Disable-vCPUHotAdd
Function Disable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End Disable-vCPUHotAdd
# Begin Get-ADDirectReports
Function Get-ADDirectReports
{
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
							Get-ADUser -Identity $PSItem -Properties mail, manager | Select-Object -Property Name, SamAccountName, Mail, @{ Name = "Manager"; Expression = { (Get-Aduser -identity $psitem.manager).samaccountname } }
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
					Get-ADUser -Properties mail, manager | Select-Object -Property Name, SamAccountName, Mail, @{ Name = "Manager"; Expression = { (Get-Aduser -identity $psitem.manager).samaccountname } }
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

<#
# Find all direct user reporting to Test_director
Get-ADDirectReports -Identity Test_director

# Find all Indirect user reporting to Test_director
Get-ADDirectReports -Identity Test_director -Recurse
#>
# End Get-ADDirectReports
# Begin Get-ADUserBadPasswords
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
# End Get-ADUserBadPasswords
# Begin Clean-Memory
Function Clean-Memory {
Get-Variable |
 Where-Object { $startupVariables -notcontains $_.Name } |
 ForEach-Object {
  try { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue}
  catch { }
 }
}
# End Clean-Memory
# Begin Remove-UserVariables
Function Remove-UserVariable {
    [CmdletBinding(SupportsShouldProcess)]
    param ()
    if ($StartupVars) {
        $UserVars = Get-Variable -Exclude $StartupVars -Scope Global
        foreach ($var in $UserVars) {
            try {
                Remove-Variable -Name $var.Name -Force -Scope Global -ErrorAction Stop
                Write-Verbose -Message "Variable '$($var.Name)' has been removed."
            }
            catch {Write-Warning -Message "An error has occured. Error Details: $($_.Exception.Message)"}           
        }
    } else {Write-Warning -Message '$StartupVars has not been added to your PowerShell profile'}    
}

$StartupVars = @()
$StartupVars = Get-Variable | Select-Object -ExpandProperty Name
# End Remove-UserVariable
# Begin Enable-PSTranscriptionLogging
Function Enable-PSTranscriptionLogging {
	param(
		[Parameter(Mandatory)]
		[string]$OutputDirectory
	)

     # Registry path
     $basePath = 'HKLM:\SOFTWARE\WOW6432Node\Policies\Microsoft\Windows\PowerShell\Transcription'

     # Create the key if it does not exist
     if(-not (Test-Path $basePath))
     {
         $null = New-Item $basePath -Force

         # Create the correct properties
         New-ItemProperty $basePath -Name "EnableInvocationHeader" -PropertyType Dword
         New-ItemProperty $basePath -Name "EnableTranscripting" -PropertyType Dword
         New-ItemProperty $basePath -Name "OutputDirectory" -PropertyType String
     }

     # These can be enabled (1) or disabled (0) by changing the value
     Set-ItemProperty $basePath -Name "EnableInvocationHeader" -Value "1"
     Set-ItemProperty $basePath -Name "EnableTranscripting" -Value "1"
     Set-ItemProperty $basePath -Name "OutputDirectory" -Value $OutputDirectory

}
# End Enable-PSTranscriptionLogging
# Begin Get-VMOSList

Function Get-VMOSList {
    [cmdletbinding()]
    param($vCenter)
    
    Connect-VIServer $vCenter  | Out-Null
    
    [array]$osNameObject       = $null
    $vmHosts                   = Get-VMHost
    $i = 0
    
    foreach ($h in $vmHosts) {
        
        Write-Progress -Activity "Going through each host in $vCenter..." -Status "Current Host: $h" -PercentComplete ($i/$vmHosts.Count*100)
        $osName = ($h | Get-VM | Get-View).Summary.Config.GuestFullName
        [array]$guestOSList += $osName
        Write-Verbose "Found OS: $osName"
        
        $i++    
 
    
    }
    
    $names = $guestOSList | Select-Object -Unique
    
    $i = 0
    
    foreach ($n in $names) { 
    
        Write-Progress -Activity "Going through VM OS Types in $vCenter..." -Status "Current Name: $n" -PercentComplete ($i/$names.Count*100)
        $vmTotal = ($guestOSList | ?{$_ -eq $n}).Count
        
        $osNameProperty  = @{'Name'=$n} 
        $osNameProperty += @{'Total VMs'=$vmTotal}
        $osNameProperty += @{'vCenter'=$vcenter}
        
        $osnO             = New-Object PSObject -Property $osNameProperty
        $osNameObject     += $osnO
        
        $i++
    
    }    
    Disconnect-VIserver -force -confirm:$false
        
    Return $osNameObject
}
# End Get-VMOSList
# Begin Get-OlderFiles
Function Get-OlderFiles {

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Path
)
 
#Check and return if the provided Path not found
if(-not (Test-Path -Path $Path) ) {
    Write-Error "Provided Path ($Path) not found"
    return
}
 
try {
    $files = Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue
    foreach($file in $files) {
         
        #Skip directories as the current focus is only on files
        if($file.PSIsContainer) {
            Continue
        }
 
        $last_modified = $file.Lastwritetime
        $time_diff_in_days = [math]::floor(((get-date) - $last_modified).TotalDays)
        $obj = New-Object -TypeName PSObject
        $obj | Add-Member -MemberType NoteProperty -Name FileName -Value $file.Name
        $obj | Add-Member -MemberType NoteProperty -Name FullPath -Value $file.FullName
        $obj | Add-Member -MemberType NoteProperty -Name AgeInDays -Value $time_diff_in_days
        $obj | Add-Member -MemberType NoteProperty -Name SizeInMB -Value $([Math]::Round(($file.Length / 1MB),3))
        $obj
    }
} catch {
    Write-Error "Error occurred. $_"
}}
#End Get-OlderFiles
#Begin Find-User
Function Find-User ($username) {
  $homeserver = ((get-aduser -id $username -prop homedirectory).Homedirectory -split "\\")[2]
  $query = "SELECT UserName,ComputerName,ActiveTime,IdleTime from win32_serversession WHERE UserName like '$username'"
  $results = Get-WmiObject -Namespace root\cimv2 -computer $homeServer -Query $query | Select UserName,ComputerName,ActiveTime,IdleTime
  foreach ($result in $results) {
    $hostname = ""
    $hostname = [System.net.Dns]::GetHostEntry($result.ComputerName).hostname
    $result | Add-Member -Type NoteProperty -Name HostName -Value $hostname -force
    $result | Add-Member -Type NoteProperty -Name HomeServer -Value $homeServer -force
  }
  $results
}

# Find one or more users
#$users = "user1", "user2", "user3"
#$users | % {Find-User $_} | ft -wrap -auto

# Find the members of a group
#get-adgroupmember -id SG-Group1 | % {Find-User $_.samaccountname} | ft -wrap -auto
#End Find-User
#Begin Force-WSUSChecking
Function Force-WSUSCheckin($Computer)
{
   Invoke-Command -computername $Computer -scriptblock { Start-Service wuauserv -Verbose }
   # Have to use psexec with the -s parameter as otherwise we receive an "Access denied" message loading the comobject
   $Cmd = '$updateSession = new-object -com "Microsoft.Update.Session";$updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates'
   psexec -s \\$Computer powershell.exe -command $Cmd
   Write-host "Waiting 10 seconds for SyncUpdates webservice to complete to add to the wuauserv queue so that it can be reported on"
   Start-sleep -seconds 10
   Invoke-Command -computername $Computer -scriptblock
   {
      # Now that the system is told it CAN report in, run every permutation of commands to actually trigger the report in operation
      wuauclt /detectnow
      (New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow()
      wuauclt /reportnow
      c:\windows\system32\UsoClient.exe startscan
   }
}
#End Force-WSUSChecking
#Begin Get-EffectiveAccess
Function Get-EffectiveAccess {
[CmdletBinding()]
param(
    [Parameter(
        Mandatory,
        ValueFromPipelineByPropertyName
    )]
    [ValidatePattern(
        '(?:(CN=([^,]*)),)?(?:((?:(?:CN|OU)=[^,]+,?)+),)?((?:DC=[^,]+,?)+)$'
    )][string]$DistinguishedName,
    [switch]$IncludeOrphan
)

    begin
    {
        # requires -Modules ActiveDirectory
        $ErrorActionPreference = 'Stop'
        $GUIDMap = @{}
        $domain = Get-ADRootDSE
        $z = '00000000-0000-0000-0000-000000000000'
        $hash = @{
            SearchBase = $domain.schemaNamingContext
            LDAPFilter = '(schemaIDGUID=*)'
            Properties = 'name','schemaIDGUID'
            ErrorAction = 'SilentlyContinue'
        }
        $schemaIDs = Get-ADObject @hash 

        $hash = @{
            SearchBase = "CN=Extended-Rights,$($domain.configurationNamingContext)"
            LDAPFilter = '(objectClass=controlAccessRight)'
            Properties = 'name','rightsGUID'
            ErrorAction = 'SilentlyContinue'
        }
        $extendedRigths = Get-ADObject @hash

        foreach($i in $schemaIDs)
        {
            if(-not $GUIDMap.ContainsKey([System.GUID]$i.schemaIDGUID))
            {
                $GUIDMap.add([System.GUID]$i.schemaIDGUID,$i.name)
            }
        }
        foreach($i in $extendedRigths)
        {
            if(-not $GUIDMap.ContainsKey([System.GUID]$i.rightsGUID))
            {
                $GUIDMap.add([System.GUID]$i.rightsGUID,$i.name)
            }
        }
    }

    process
    {
        $result = [system.collections.generic.list[pscustomobject]]::new()
        $object = Get-ADObject $DistinguishedName
        $acls = (Get-ACL "AD:$object").Access
        
        foreach($acl in $acls)
        {
            
            $objectType = if($acl.ObjectType -eq $z)
            {
                'All Objects (Full Control)'
            }
            else
            {
                $GUIDMap[$acl.ObjectType]
            }

            $inheritedObjType = if($acl.InheritedObjectType -eq $z)
            {
                'Applied to Any Inherited Object'
            }
            else
            {
                $GUIDMap[$acl.InheritedObjectType]
            }

            $result.Add(
                [PSCustomObject]@{
                    Name = $object.Name
                    IdentityReference = $acl.IdentityReference
                    AccessControlType = $acl.AccessControlType
                    ActiveDirectoryRights = $acl.ActiveDirectoryRights
                    ObjectType = $objectType
                    InheritedObjectType = $inheritedObjType
                    InheritanceType = $acl.InheritanceType
                    IsInherited = $acl.IsInherited
            })
        }
        
        if(-not $IncludeOrphan.IsPresent)
        {
            $result | Sort-Object IdentityReference |
            Where-Object {$_.IdentityReference -notmatch 'S-1-*'}
            return
        }

        return $result | Sort-Object IdentityReference
    }
}
#End Get-EffectiveAccess
#Begin Get-InstalledApplication
Function Get-InstalledApplication {
  [CmdletBinding()]
  Param(
    [Parameter(
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true
    )]
    [String[]]$ComputerName=$ENV:COMPUTERNAME,

    [Parameter(Position=1)]
    [String[]]$Properties,

    [Parameter(Position=2)]
    [String]$IdentifyingNumber,

    [Parameter(Position=3)]
    [String]$Name,

    [Parameter(Position=4)]
    [String]$Publisher
  )
  Begin{
    Function IsCpuX86 ([Microsoft.Win32.RegistryKey]$hklmHive){
      $regPath='SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
      $key=$hklmHive.OpenSubKey($regPath)

      $cpuArch=$key.GetValue('PROCESSOR_ARCHITECTURE')

      if($cpuArch -eq 'x86'){
        return $true
      }else{
        return $false
      }
    }
  }
  Process{
    foreach($computer in $computerName){
      $regPath = @(
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
      )

      Try{
        $hive=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
          [Microsoft.Win32.RegistryHive]::LocalMachine, 
          $computer
        )
        if(!$hive){
          continue
        }
        
        # if CPU is x86 do not query for Wow6432Node
        if($IsCpuX86){
          $regPath=$regPath[0]
        }

        foreach($path in $regPath){
          $key=$hive.OpenSubKey($path)
          if(!$key){
            continue
          }
          foreach($subKey in $key.GetSubKeyNames()){
            $subKeyObj=$null
            if($PSBoundParameters.ContainsKey('IdentifyingNumber')){
              if($subKey -ne $IdentifyingNumber -and 
                $subkey.TrimStart('{').TrimEnd('}') -ne $IdentifyingNumber){
                continue
              }
            }
            $subKeyObj=$key.OpenSubKey($subKey)
            if(!$subKeyObj){
              continue
            }
            $outHash=New-Object -TypeName Collections.Hashtable
            $appName=[String]::Empty
            $appName=($subKeyObj.GetValue('DisplayName'))
            if($PSBoundParameters.ContainsKey('Name')){
              if($appName -notlike $name){
                continue
              }
            }
            if($appName){
              if($PSBoundParameters.ContainsKey('Properties')){
                if($Properties -eq '*'){
                  foreach($keyName in ($hive.OpenSubKey("$path\$subKey")).GetValueNames()){
                    Try{
                      $value=$subKeyObj.GetValue($keyName)
                      if($value){
                        $outHash.$keyName=$value
                      }
                    }Catch{
                      Write-Warning "Subkey: [$subkey]: $($_.Exception.Message)"
                      continue
                    }
                  }
                }else{
                  foreach ($prop in $Properties){
                    $outHash.$prop=($hive.OpenSubKey("$path\$subKey")).GetValue($prop)
                  }
                }
              }
              $outHash.Name=$appName
              $outHash.IdentifyingNumber=$subKey
              $outHash.Publisher=$subKeyObj.GetValue('Publisher')
              if($PSBoundParameters.ContainsKey('Publisher')){
                if($outHash.Publisher -notlike $Publisher){
                  continue
                }
              }
              $outHash.ComputerName=$computer
              $outHash.Path=$subKeyObj.ToString()
              New-Object -TypeName PSObject -Property $outHash
            }
          }
        }
      }Catch{
        Write-Error $_
      }
    }
  }
  End{}
}
#End Get-InstalledApplication
#Begin Get-IPv6InWindows
Function Get-IPv6InWindows
{
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
#End Get-IPv6InWindows
#Begin Get-Uptime
Function Get-Uptime {
<#
.Synopsis
    This will check how long the computer has been running and when was it last rebooted.
    For updated help and examples refer to -Online version.
 
 
.NOTES
    Name: Get-Uptime
    Author: theSysadminChannel
    Version: 1.0
    DateCreated: 2018-Jun-16
 
.LINK
    https://thesysadminchannel.com/get-uptime-last-reboot-status-multiple-computers-powershell/ -
 
 
.PARAMETER ComputerName
    By default it will check the local computer.
 
 
    .EXAMPLE
    Get-Uptime -ComputerName PAC-DC01, PAC-WIN1001
 
    Description:
    Check the computers PAC-DC01 and PAC-WIN1001 and see how long the systems have been running for.
 
#>
 
    [CmdletBinding()]
    Param (
        [Parameter(
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
 
        [string[]]
            $ComputerName = $env:COMPUTERNAME
    )
 
    BEGIN {}
 
    PROCESS {
        Foreach ($Computer in $ComputerName) {
            $Computer = $Computer.ToUpper()
            Try {
                $OS = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop
                $Uptime = (Get-Date) - $OS.ConvertToDateTime($OS.LastBootUpTime)
                [PSCustomObject]@{
                    ComputerName  = $Computer
                    LastBoot      = $OS.ConvertToDateTime($OS.LastBootUpTime)
                    Uptime        = ([String]$Uptime.Days + " Days " + $Uptime.Hours + " Hours " + $Uptime.Minutes + " Minutes")
                }
 
            } catch {
                [PSCustomObject]@{
                    ComputerName  = $Computer
                    LastBoot      = "Unable to Connect"
                    Uptime        = $_.Exception.Message.Split('.')[0]
                }
 
            } finally {
                $null = $OS
                $null = $Uptime
            }
        }
    }
 
    END {}
 
}
#End Get-Uptime
#Begin Get-VMEvcMode
Function Get-VMEvcMode {
<#  
.SYNOPSIS  
    Gathers information on the EVC status of a VM
.DESCRIPTION 
    Will provide the EVC status for the specified VM
.NOTES  
    Author:  Kyle Ruddy, @kmruddy, thatcouldbeaproblem.com
.PARAMETER Name
    VM name which the Function should be ran against
.EXAMPLE
	Get-VMEvcMode -Name vmName
	Retreives the EVC status of the provided VM 
#>
[CmdletBinding()] 
	param(
	[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        $Name
  	)

    Process {
        $evVM = @()

        if ($name -is [string]) {$evVM += Get-VM -Name $Name -ErrorAction SilentlyContinue}
        elseif ($name -is [array]) {

            if ($name[0] -is [string]) {
                $name | foreach {
                    $evVM += Get-VM -Name $_ -ErrorAction SilentlyContinue
                }
            }
            elseif ($name[0] -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM = $name}

        }
        elseif ($name -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM += $name}
        
        if ($evVM -eq $null) {Write-Warning "No VMs found."}
        else {
            $output = @()
            foreach ($v in $evVM) {

                $report = "" | select Name,EVCMode
                $report.Name = $v.Name
                $report.EVCMode = $v.ExtensionData.Runtime.MinRequiredEVCModeKey
                $output += $report

            }

        return $output

        }

    }

}
#End Get-VMEVCMode
#Begin Remove-VMEVCMode
Function Remove-VMEvcMode {
<#  
.SYNOPSIS  
    Removes the EVC status of a VM
.DESCRIPTION 
    Will remove the EVC status for the specified VM
.NOTES  
    Author:  Kyle Ruddy, @kmruddy, thatcouldbeaproblem.com
.PARAMETER Name
    VM name which the Function should be ran against
.EXAMPLE
	Remove-VMEvcMode -Name vmName
	Removes the EVC status of the provided VM 
#>
[CmdletBinding()] 
	param(
	[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        $Name
  	)

    Process {
        $evVM = @()
        $updateVM = @()

        if ($name -is [string]) {$evVM += Get-VM -Name $Name -ErrorAction SilentlyContinue}
        elseif ($name -is [array]) {

            if ($name[0] -is [string]) {
                $name | foreach {
                    $evVM += Get-VM -Name $_ -ErrorAction SilentlyContinue
                }
            }
            elseif ($name[0] -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM = $name}

        }
        elseif ($name -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM += $name}
        
        if ($evVM -eq $null) {Write-Warning "No VMs found."}
        else {
            foreach ($v in $evVM) {

                if (($v.HardwareVersion -ge 'vmx-14' -and $v.PowerState -eq 'PoweredOff') -or ($v.Version -ge 'v14' -and $v.PowerState -eq 'PoweredOff')) {

                    $v.ExtensionData.ApplyEvcModeVM_Task($null, $true) | Out-Null
                    $updateVM += $v.Name
                                        
                }
                else {Write-Warning $v.Name + " does not have the minimum requirements of being Hardware Version 14 and powered off."}

            }

            if ($updateVM) {
            
            Start-Sleep -Seconds 2
            Get-VMEvcMode -Name $updateVM
            
            }

        }

    }

}
#End Remove-VMEvcMode
#Begin Set-VMEvcMode
Function Set-VMEvcMode {
<#  
.SYNOPSIS  
    Configures the EVC status of a VM
.DESCRIPTION 
    Will configure the EVC status for the specified VM
.NOTES  
    Author:  Kyle Ruddy, @kmruddy, thatcouldbeaproblem.com
.PARAMETER Name
    VM name which the Function should be ran against
.PARAMETER EvcMode
    The EVC Mode key which should be set
.EXAMPLE
	Set-VMEvcMode -Name vmName -EvcMode intel-sandybridge
	Configures the EVC status of the provided VM to be 'intel-sandybridge'
#>
[CmdletBinding()] 
	param(
	[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        $Name,
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateSet("intel-merom","intel-penryn","intel-nehalem","intel-westmere","intel-sandybridge","intel-ivybridge","intel-haswell","intel-broadwell","intel-skylake","amd-rev-e","amd-rev-f","amd-greyhound-no3dnow","amd-greyhound","amd-bulldozer","amd-piledriver","amd-steamroller","amd-zen")]
        $EvcMode
  	)

    Process {
        $evVM = @()
        $updateVM = @()

        if ($name -is [string]) {$evVM += Get-VM -Name $Name -ErrorAction SilentlyContinue}
        elseif ($name -is [array]) {

            if ($name[0] -is [string]) {
                $name | foreach {
                    $evVM += Get-VM -Name $_ -ErrorAction SilentlyContinue
                }
            }
            elseif ($name[0] -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM = $name}

        }
        elseif ($name -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM += $name}
        
        if ($evVM -eq $null) {Write-Warning "No VMs found."}
        else {

            $si = Get-View ServiceInstance
            $evcMask = $si.Capability.SupportedEvcMode | where-object {$_.key -eq $EvcMode} | select -ExpandProperty FeatureMask

            foreach ($v in $evVM) {

                if (($v.HardwareVersion -ge 'vmx-14' -and $v.PowerState -eq 'PoweredOff') -or ($v.Version -ge 'v14' -and $v.PowerState -eq 'PoweredOff')) {

                    $v.ExtensionData.ApplyEvcModeVM_Task($evcMask, $true) | Out-Null
                    $updateVM += $v.Name
                                        
                }
                else {Write-Warning $v.Name + " does not have the minimum requirements of being Hardware Version 14 and powered off."}

            }

            if ($updateVM) {
            
            Start-Sleep -Seconds 2
            Get-VMEvcMode -Name $updateVM
            
            }

        }

    }

}
#End Set-VMEvcMode
#Begin Uninstall-Modules
Function UnInstall-Modules {

[CmdletBinding()]
param(
[Parameter(Mandatory = $false)]
[string]
$RetentionMonths = 3
)

if ($PSCmdlet.MyInvocation.BoundParameters[“Debug”].IsPresent) {
$DebugPreference = “Continue”
}
$CMDLetName = $MyInvocation.MyCommand.Name

# Get a list of the current modules installed.
Write-Debug “Getting list of current modules installed …”
$Modules = Get-InstalledModule
$Counter = 0 # Used to track count of un-installations

foreach ($Module in $Modules) {
Write-Debug $($Module.Name) # List out all the modules installed.
}

foreach ($Module in $Modules) {

Write-Host “`n”
$ModuleVersions = Get-InstalledModule -Name $($Module.Name) -AllVersions # Get all versions of the module
$ModuleVersionsArray = New-Object System.Collections.ArrayList
foreach ($ModuleVersion in $ModuleVersions) {
$ModuleVersionsArray.Add($ModuleVersion.Version) > $Null
}
Write-Debug “Reviewing module: $($Module.name) – Versions installed: $($ModuleVersionsArray.Count)”

$VersionsToKeepArray = New-Object System.Collections.ArrayList
$MajorVersions = @($ModuleVersionsArray.Major | Get-Unique) # Get unique majors
$MinorVersions = @($ModuleVersionsArray.Minor | Get-Unique) # Get unique minors

foreach ($MajorVersion in $MajorVersions) {
foreach ($MinorVersion in $MinorVersions) {
$ReturnedVersion = (Get-InstalledModule -Name $($Module.Name) -MaximumVersion “${MajorVersion}.${MinorVersion}.99999” -ErrorAction SilentlyContinue)
$VersionsToKeepArray.add($ReturnedVersion) > $Null # Versions to keep
$ModuleVersionsArray.Remove($ReturnedVersion.Version) # Remove versions we’re keeping.
}
}

# Groom the builds
if ($ModuleVersionsArray) {
foreach ($Version in $ModuleVersionsArray) {
Write-Debug “Removing Module: $($Module.Name) – Version: ${Version} ”
try {
Uninstall-Module -Name $($Module.Name) -RequiredVersion “${Version}” -ErrorAction Stop
$Counter++
}
catch {
Write-Warning “Problem”
}
}
}
else {
Write-Debug “No builds to remove”
}

# Evaluate removing previous builds older than retention period.
$VersionsToRemoveArray = New-Object System.Collections.ArrayList # Create an array a versions to remove
$Oldest = ($VersionsToKeepArray.version | Measure-Object -Minimum).Minimum # Get oldest version
$Newest = ($VersionsToKeepArray.version | Measure-Object -Maximum).Maximum # Get newest version
$ReturnedVersion = (Get-InstalledModule -Name $($Module.Name) -RequiredVersion $Oldest) # Find the oldest of the keepers
if ($Oldest -ne $Newest) {
# Skip adding it the current is both newest and oldest.
$VersionsToRemoveArray.add($ReturnedVersion) > $Null # Versions to remove
}

if ($VersionsToRemoveArray) {
foreach ($Module in $VersionsToRemoveArray) {
if ($Module.version -eq $Oldest -and $Module.InstalledDate -lt (get-date).AddMonths( – ${RetentionMonths})) {
try {
Uninstall-Module -Name $($Module.Name) -RequiredVersion “${Version}” -ErrorAction Stop
$Counter++
}
catch {
Write-Warning “Problem”
}
}
else {
Write-Debug “Module: $($Module.Name) – Version: $($Module.version) is not yet older than retention of ${RetentionMonths} months, skipping removal. ”
}
}
}

} # For each module end

if ($Counter -gt 0) {
Write-Debug “Removed ${Counter} module versions”
}
} # Function end
#End Uninstall-Modules
#Begin VMWareFunctions
# Enable or Disable Hot Add Memory/CPU
# Enable-MemHotAdd $ServerName
# Disable-MemHotAdd $ServerName
# Enable-vCPUHotAdd $ServerName
# Disable-vCPUHotAdd $ServerName




Function Enable-MemHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-MemHotAdd($vm){
$vmview = Get-VM $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Enable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Get-VMEVCMode {
    <#  .Description
        Code to get VMs' EVC mode and that of the cluster in which the VMs reside.  May 2014, vNugglets.com
        .Example
        Get-VMEVCMode -Cluster myCluster | ?{$_.VMEVCMode -ne $_.ClusterEVCMode}
        Get all VMs in given clusters and return, for each, an object with the VM's- and its cluster's EVC mode, if any
        .Outputs
        PSCustomObject
    #>
    param(
        ## Cluster name pattern (regex) to use for getting the clusters whose VMs to get
        [string]$Cluster_str = ".+"
    )
 
    process {
        ## get the matching cluster View objects
        Get-View -ViewType ClusterComputeResource -Property Name,Summary -Filter @{"Name" = $Cluster_str} | Foreach-Object {
            $viewThisCluster = $_
            ## get the VMs Views in this cluster
            Get-View -ViewType VirtualMachine -Property Name,Runtime.PowerState,Summary.Runtime.MinRequiredEVCModeKey -SearchRoot $viewThisCluster.MoRef | Foreach-Object {
                ## create new PSObject with some nice info
                New-Object -Type PSObject -Property ([ordered]@{
                    Name = $_.Name
                    PowerState = $_.Runtime.PowerState
                    VMEVCMode = $_.Summary.Runtime.MinRequiredEVCModeKey
                    ClusterEVCMode = $viewThisCluster.Summary.CurrentEVCModeKey
                    ClusterName = $viewThisCluster.Name
                })
            } ## end foreach-object
        } ## end foreach-object
    } ## end process
} ## end Function
#End VMWare-Functions
#Begin SEPVersion Check
Function Get-SEPVersion {

[CmdletBinding()]
param(
[Parameter(Position=0,Mandatory=$true,HelpMessage='Name of the computer to query SEP for',
ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[Alias('CN','__SERVER','IPAddress','Server')]
[System.String]
$ComputerName
)

begin {
# Create object to enable access to the months of the year
$DateTimeFormat = New-Object -TypeName System.Globalization.DateTimeFormatInfo

# Set Registry keys to query

If((Get-WmiObject -ComputerName $ComputerName -Class Win32_OperatingSystem).OSArchitecture -eq '32-bit')
{
$SMCKey = 'SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC'
$AVKey = 'SOFTWARE\\Symantec\\Symantec Endpoint Protection\\AV'
$SylinkKey = 'SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink'
}
Else
{
$SMCKey = 'SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC' 
$AVKey = 'SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV' 
$SylinkKey = 'SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink' 
}
    }


process {


try {

# Connect to Registry
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ComputerName)

# Obtain Product Version value
$SMCRegKey = $reg.opensubkey($SMCKey)
$SEPVersion = $SMCRegKey.GetValue('ProductVersion')

# Obtain Pattern File Date Value
$AVRegKey = $reg.opensubkey($AVKey)
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate')

# Convert PatternFileDate to readable date
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970)
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]]
$AVDayFileDate = [string]$AVPatternFileDate[2]
$AVFileVersionDate = $AVDayFileDate + ' ' + $AVMonthFileDate + ' ' + $AVYearFileDate

# Obtain Sylink Group value
$SylinkRegKey = $reg.opensubkey($SylinkKey)
$SylinkGroup = $SylinkRegKey.GetValue('CurrentGroup')

}

catch [System.Management.Automation.MethodInvocationException]

{
$SEPVersion = 'Unable to connect to computer'
$AVFileVersionDate = ''
$SylinkGroup = ''
}

$MYObject = '' | Select-Object -Property ComputerName,SEPProductVersion,SEPDefinitionDate,SylinkGroup
$MYObject.ComputerName = $ComputerName
$MYObject.SEPProductVersion = $SEPVersion
$MYObject.SEPDefinitionDate = $AVFileVersionDate
$MYObject.SylinkGroup = $SylinkGroup
$MYObject

}
}
# End SEPVersion Check
# Begin Set-DnsServerIpAddress
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
# End Set-DnsServerIpAddress
# Begin Get-TaskPlus
Function Get-TaskPlus {
<#  
.SYNOPSIS  Returns vSphere Task information   
.DESCRIPTION The Function will return vSphere task info. The
available parameters allow server-side filtering of the
results
.NOTES  Author:  Luc Dekens  
.PARAMETER Alarm
When specified the Function returns tasks triggered by
specified alarm
.PARAMETER Entity
When specified the Function returns tasks for the
specific vSphere entity
.PARAMETER Recurse
Is used with the Entity. The Function returns tasks
for the Entity and all it's children
.PARAMETER State
Specify the State of the tasks to be returned. Valid
values are: error, queued, running and success
.PARAMETER Start
The start date of the tasks to retrieve
.PARAMETER Finish
The end date of the tasks to retrieve.
.PARAMETER UserName
Only return tasks that were started by a specific user
.PARAMETER MaxSamples
Specify the maximum number of tasks to return
.PARAMETER Reverse
When true, the tasks are returned newest to oldest. The
default is oldest to newest
.PARAMETER Server
The vCenter instance(s) for which the tasks should
be returned
.PARAMETER Realtime
A switch, when true the most recent tasks are also returned.
.PARAMETER Details
A switch, when true more task details are returned
.PARAMETER Keys
A switch, when true all the keys are returned
.EXAMPLE
PS> Get-TaskPlus -Start (Get-Date).AddDays(-1)
.EXAMPLE
PS> Get-TaskPlus -Alarm $alarm -Details
#>
param(
[CmdletBinding()]
[VMware.VimAutomation.ViCore.Impl.V1.Alarm.AlarmDefinitionImpl]$Alarm,
[VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Entity,
[switch]$Recurse = $false,
[VMware.Vim.TaskInfoState[]]$State,
[DateTime]$Start,
[DateTime]$Finish,
[string]$UserName,
[int]$MaxSamples = 100000,
[switch]$Reverse = $true,
[VMware.VimAutomation.ViCore.Impl.V1.VIServerImpl[]]$Server = $global:DefaultVIServer,
[switch]$Realtime,
[switch]$Details,
[switch]$Keys,
[int]$WindowSize = 1000
)
begin {
Function Get-TaskDetails {
param(
[VMware.Vim.TaskInfo[]]$Tasks
)
begin {
$psV3 = $PSversionTable.PSVersion.Major -ge 3
}
process {
$tasks | ForEach-Object {
if ($psV3) {
$object = [ordered]@{ }
}
else {
$object = @{ }
}
$object.Add("Name", $_.Name)
$object.Add("Description", $_.Description.Message)
if ($Details) { $object.Add("DescriptionId", $_.DescriptionId) }
if ($Details) { $object.Add("Task Created", $_.QueueTime) }
$object.Add("Task Started", $_.StartTime)
if ($Details) { $object.Add("Task Ended", $_.CompleteTime) }
$object.Add("State", $_.State)
$object.Add("Result", $_.Result)
$object.Add("Entity", $_.EntityName)
$object.Add("VIServer", $VIObject.Name)
$object.Add("Error", $_.Error.ocalizedMessage)
if ($Details) {
$object.Add("Cancelled", (& { if ($_.Cancelled) { "Y" }else { "N" } }))
$object.Add("Reason", $_.Reason.GetType().Name.Replace("TaskReason", ""))
$object.Add("AlarmName", $_.Reason.AlarmName)
$object.Add("AlarmEntity", $_.Reason.EntityName)
$object.Add("ScheduleName", $_.Reason.Name)
$object.Add("User", $_.Reason.UserName)
}
if ($keys) {
$object.Add("Key", $_.Key)
$object.Add("ParentKey", $_.ParentTaskKey)
$object.Add("RootKey", $_.RootTaskKey)
}
New-Object PSObject -Property $object
}
}
}
$filter = New-Object VMware.Vim.TaskFilterSpec
if ($Alarm) {
$filter.Alarm = $Alarm.ExtensionData.MoRef
}
if ($Entity) {
$filter.Entity = New-Object VMware.Vim.TaskFilterSpecByEntity
$filter.Entity.entity = $Entity.ExtensionData.MoRef
if ($Recurse) {
$filter.Entity.Recursion = [VMware.Vim.TaskFilterSpecRecursionOption]::all
}
else {
$filter.Entity.Recursion = [VMware.Vim.TaskFilterSpecRecursionOption]::self
}
}
if ($State) {
$filter.State = $State
}
if ($Start -or $Finish) {
$filter.Time = New-Object VMware.Vim.TaskFilterSpecByTime
$filter.Time.beginTime = $Start
$filter.Time.endTime = $Finish
$filter.Time.timeType = [vmware.vim.taskfilterspectimeoption]::startedTime
}
if ($UserName) {
$userNameFilterSpec = New-Object VMware.Vim.TaskFilterSpecByUserName
$userNameFilterSpec.UserList = $UserName
$filter.UserName = $userNameFilterSpec
}
$nrTasks = 0
}
process {
foreach ($viObject in $Server) {
$si = Get-View ServiceInstance -Server $viObject
$tskMgr = Get-View $si.Content.TaskManager -Server $viObject 
if ($Realtime -and $tskMgr.recentTask) {
$tasks = Get-View $tskMgr.recentTask
$selectNr = [Math]::Min($tasks.Count, $MaxSamples - $nrTasks)
Get-TaskDetails -Tasks[0..($selectNr - 1)]
$nrTasks += $selectNr
}
$tCollector = Get-View ($tskMgr.CreateCollectorForTasks($filter))
if ($Reverse) {
$tCollector.ResetCollector()
$taskReadOp = $tCollector.ReadPreviousTasks
}
else {
$taskReadOp = $tCollector.ReadNextTasks
}
do {
$tasks = $taskReadOp.Invoke($WindowSize)
if (!$tasks) { break }
$selectNr = [Math]::Min($tasks.Count, $MaxSamples - $nrTasks)
Get-TaskDetails -Tasks $tasks[0..($selectNr - 1)]
$nrTasks += $selectNr
}while ($nrTasks -lt $MaxSamples)
$tCollector.DestroyCollector()
}
}
}
# End Get-TaskPlus
# Begin Get-NetworkLevelAuthentication
Function Get-NetworkLevelAuthentication
{
<#
	.SYNOPSIS
		This Function will get the NLA setting on a local machine or remote machine

	.DESCRIPTION
		This Function will get the NLA setting on a local machine or remote machine

	.PARAMETER  ComputerName
		Specify one or more computer to query

	.PARAMETER  Credential
		Specify the alternative credential to use. By default it will use the current one.
	
	.EXAMPLE
		Get-NetworkLevelAuthentication
		
		This will get the NLA setting on the localhost
	
		ComputerName     : XAVIERDESKTOP
		NLAEnabled       : True
		TerminalName     : RDP-Tcp
		TerminalProtocol : Microsoft RDP 8.0
		Transport        : tcp	

    .EXAMPLE
		Get-NetworkLevelAuthentication -ComputerName DC01
		
		This will get the NLA setting on the server DC01
	
		ComputerName     : DC01
		NLAEnabled       : True
		TerminalName     : RDP-Tcp
		TerminalProtocol : Microsoft RDP 8.0
		Transport        : tcp
	
	.EXAMPLE
		Get-NetworkLevelAuthentication -ComputerName DC01, SERVER01 -verbose
	
	.EXAMPLE
		Get-Content .\Computers.txt | Get-NetworkLevelAuthentication -verbose
		
	.NOTES
		DATE	: 2014/04/01
		AUTHOR	: Francois-Xavier Cat
		WWW		: http://lazywinadmin.com
		Twitter	: @lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the Function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				[pscustomobject][ordered]@{
					'ComputerName' = $NLAinfo.PSComputerName
					'NLAEnabled' = $NLAinfo.UserAuthenticationRequired -as [bool]
					'TerminalName' = $NLAinfo.TerminalName
					'TerminalProtocol' = $NLAinfo.TerminalProtocol
					'Transport' = $NLAinfo.transport
				}
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{
		
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}


Function Set-NetworkLevelAuthentication
{
<#
	.SYNOPSIS
		This Function will set the NLA setting on a local machine or remote machine

	.DESCRIPTION
		This Function will set the NLA setting on a local machine or remote machine

	.PARAMETER  ComputerName
		Specify one or more computers
	
	.PARAMETER EnableNLA
		Specify if the NetworkLevelAuthentication need to be set to $true or $false
	
	.PARAMETER  Credential
		Specify the alternative credential to use. By default it will use the current one.

	.EXAMPLE
		Set-NetworkLevelAuthentication -EnableNLA $true

		ReturnValue                             PSComputerName                         
		-----------                             --------------                         
		                                        XAVIERDESKTOP      
	
	.NOTES
		DATE	: 2014/04/01
		AUTHOR	: Francois-Xavier Cat
		WWW		: http://lazywinadmin.com
		Twitter	: @lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Parameter(Mandatory)]
		[Bool]$EnableNLA,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
				
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the Function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				$NLAinfo | Invoke-CimMethod -MethodName SetUserAuthenticationRequired -Arguments @{ UserAuthenticationRequired = $EnableNLA } -ErrorAction 'Continue' -ErrorVariable ErrorProcessInvokeWmiMethod
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
				if ($ErrorProcessInvokeWmiMethod) { Write-Warning -Message "PROCESS Error - $ErrorProcessInvokeWmiMethod" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{	
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}
# End Get-NetworkLevelAuthentication
#Begin Export-Xlsx
Function Export-Xlsx {
<#
.SYNOPSIS
Exports data to an Excel workbook
.DESCRIPTION
Exports data to an Excel workbook and applies cosmetics. 
Optionally add a title, autofilter, autofit and a chart.
Allows for export to .xls and .xlsx format. If .xlsx is
specified but not available (Excel 2003) the data will
be exported to .xls.
.NOTES
Author:  Gilbert van Griensven
Based on
https://www.lucd.info/2010/05/29/beyond-export-csv-export-xls/
.PARAMETER InputData
The data to be exported to Excel
.PARAMETER Path
The path of the Excel file. 
Defaults to %HomeDrive%\Export.xlsx.
.PARAMETER WorksheetName
The name of the worksheet. Defaults to filename
in $Path without extension.
.PARAMETER ChartType
Name of an Excel chart to be added.
.PARAMETER Title
Adds a title to the worksheet.
.PARAMETER SheetPosition
Adds the worksheet either to the 'begin' or 'end' of
the Excel file. This parameter is ignored when creating
a new Excel file.
.PARAMETER ChartOnNewSheet
Adds a chart to a new worksheet instead of to the
worksheet containing data. The Chart will be placed after
the sheet containing data. Only works when parameter
ChartType is used.
.PARAMETER AppendWorksheet
Appends a worksheet to an existing Excel file.
This parameter is ignored when creating a new Excel file.
.PARAMETER Borders
Adds borders to all cells. Defaults to True.
.PARAMETER HeaderColor
Applies background color to the header row. 
Defaults to True.
.PARAMETER AutoFit
Apply autofit to columns. Defaults to True.
.PARAMETER AutoFilter
Apply autofilter. Defaults to True.
.PARAMETER PassThrough
When enabled returns file object of the generated file.
.PARAMETER Force
Overwrites existing Excel sheet. When this switch is
not used but the Excel file already exists, a new file
with datestamp will be generated. This switch is ignored
when using the AppendWorksheet switch.
.EXAMPLE
Get-Process | Export-Xlsx D:\Data\ProcessList.xlsx
.EXAMPLE
Get-ADuser -Filter {enabled -ne $True} | 
Select-Object Name,Surname,GivenName,DistinguishedName | 
Export-Xlsx -Path 'D:\Data\Disabled Users.xlsx' -Title 'Disabled users of Contoso.com'
.EXAMPLE
Get-Process | Sort-Object CPU -Descending | 
Export-Xlsx -Path D:\Data\Processes_by_CPU.xlsx
.EXAMPLE
Export-Xlsx (Get-Process) -AutoFilter:$False -PassThrough |
Invoke-Item
#>
[CmdletBinding()]
Param (
[Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$True)]
[ValidateNotNullOrEmpty()]
$InputData,
[Parameter(Position=1)]
[ValidateScript({
$ReqExt = [System.IO.Path]::GetExtension($_)
(          $ReqExt -eq ".xls") -or
(          $ReqExt -eq ".xlsx")
})]
$Path = (Join-Path $env:HomeDrive "Export.xlsx"),
[Parameter(Position=2)] $WorksheetName = [System.IO.Path]::GetFileNameWithoutExtension($Path),
[Parameter(Position=3)]
[ValidateSet("xl3DArea","xl3DAreaStacked","xl3DAreaStacked100","xl3DBarClustered",
"xl3DBarStacked","xl3DBarStacked100","xl3DColumn","xl3DColumnClustered",
"xl3DColumnStacked","xl3DColumnStacked100","xl3DLine","xl3DPie",
"xl3DPieExploded","xlArea","xlAreaStacked","xlAreaStacked100",
"xlBarClustered","xlBarOfPie","xlBarStacked","xlBarStacked100",
"xlBubble","xlBubble3DEffect","xlColumnClustered","xlColumnStacked",
"xlColumnStacked100","xlConeBarClustered","xlConeBarStacked","xlConeBarStacked100",
"xlConeCol","xlConeColClustered","xlConeColStacked","xlConeColStacked100",
"xlCylinderBarClustered","xlCylinderBarStacked","xlCylinderBarStacked100","xlCylinderCol",
"xlCylinderColClustered","xlCylinderColStacked","xlCylinderColStacked100","xlDoughnut",
"xlDoughnutExploded","xlLine","xlLineMarkers","xlLineMarkersStacked",
"xlLineMarkersStacked100","xlLineStacked","xlLineStacked100","xlPie",
"xlPieExploded","xlPieOfPie","xlPyramidBarClustered","xlPyramidBarStacked",
"xlPyramidBarStacked100","xlPyramidCol","xlPyramidColClustered","xlPyramidColStacked",
"xlPyramidColStacked100","xlRadar","xlRadarFilled","xlRadarMarkers",
"xlStockHLC","xlStockOHLC","xlStockVHLC","xlStockVOHLC",
"xlSurface","xlSurfaceTopView","xlSurfaceTopViewWireframe","xlSurfaceWireframe",
"xlXYScatter","xlXYScatterLines","xlXYScatterLinesNoMarkers","xlXYScatterSmooth",
"xlXYScatterSmoothNoMarkers")]
[PSObject] $ChartType,
[Parameter(Position=4)] $Title,
[Parameter(Position=5)] [ValidateSet("begin","end")] $SheetPosition = "begin",
[Switch] $ChartOnNewSheet,
[Switch] $AppendWorksheet,
[Switch] $Borders = $True,
[Switch] $HeaderColor = $True,
[Switch] $AutoFit = $True,
[Switch] $AutoFilter = $True,
[Switch] $PassThrough,
[Switch] $Force
)
Begin {
Function Convert-NumberToA1 {
Param([parameter(Mandatory=$true)] [int]$number)
$a1Value = $null
While ($number -gt 0) {
$multiplier = [int][system.math]::Floor(($number / 26))
$charNumber = $number - ($multiplier * 26)
If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 }
$a1Value = [char]($charNumber + 96) + $a1Value
$number = $multiplier
}
Return $a1Value
}
$Script:WorkingData = @()
}
Process {
$Script:WorkingData += $InputData
}
End {
$Props = $Script:WorkingData[0].PSObject.properties | % { $_.Name }
$Rows = $Script:WorkingData.Count+1
$Cols = $Props.Count
$A1Cols = Convert-NumberToA1 $Cols
$Array = New-Object 'object[,]' $Rows,$Cols
$Col = 0
$Props | % {
$Array[0,$Col] = $_.ToString()
$Col++
}
$Row = 1
$Script:WorkingData | % {
$Item = $_
$Col = 0
$Props | % {
If ($Item.($_) -eq $Null) {
$Array[$Row,$Col] = ""
} Else {
$Array[$Row,$Col] = $Item.($_).ToString()
}
$Col++
}
$Row++
}
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $False
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookNormal
If ([System.IO.Path]::GetExtension($Path) -eq '.xlsx') {
If ($xl.Version -lt 12) {
$Path = $Path.Replace(".xlsx",".xls")
} Else {
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookDefault
}
}
If (Test-Path -Path $Path -PathType "Leaf") {
If ($AppendWorkSheet) {
$wb = $xl.Workbooks.Open($Path)
If ($SheetPosition -eq "end") {
$wb.Worksheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count)) | Out-Null
} Else {
$wb.Worksheets.Add($wb.Worksheets.Item(1)) | Out-Null
}
} Else {
If (!($Force)) {
$Path = $Path.Insert($Path.LastIndexOf(".")," - $(Get-Date -Format "ddMMyyyy-HHmm")")
}
$wb = $xl.Workbooks.Add()
While ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
}
} Else {
$wb = $xl.Workbooks.Add()
While ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
}
$ws = $wb.ActiveSheet
Try { $ws.Name = $WorksheetName }
Catch { }
If ($Title) {
$ws.Cells.Item(1,1) = $Title
$TitleRange = $ws.Range("a1","$($A1Cols)2")
$TitleRange.Font.Size = 18
$TitleRange.Font.Bold=$True
$TitleRange.Font.Name = "Cambria"
$TitleRange.Font.ThemeFont = 1
$TitleRange.Font.ThemeColor = 4
$TitleRange.Font.ColorIndex = 55
$TitleRange.Font.Color = 8210719
$TitleRange.Merge()
$TitleRange.VerticalAlignment = -4160
$usedRange = $ws.Range("a3","$($A1Cols)$($Rows + 2)")
If ($HeaderColor) {
$ws.Range("a3","$($A1Cols)3").Interior.ColorIndex = 48
$ws.Range("a3","$($A1Cols)3").Font.Bold = $True
}
} Else {
$usedRange = $ws.Range("a1","$($A1Cols)$($Rows)")
If ($HeaderColor) {
$ws.Range("a1","$($A1Cols)1").Interior.ColorIndex = 48
$ws.Range("a1","$($A1Cols)1").Font.Bold = $True
}
}
$usedRange.Value2 = $Array
If ($Borders) {
$usedRange.Borders.LineStyle = 1
$usedRange.Borders.Weight = 2
}
If ($AutoFilter) { $usedRange.AutoFilter() | Out-Null }
If ($AutoFit) { $ws.UsedRange.EntireColumn.AutoFit() | Out-Null }
If ($ChartType) {
[Microsoft.Office.Interop.Excel.XlChartType]$ChartType = $ChartType
If ($ChartOnNewSheet) {
$wb.Charts.Add().ChartType = $ChartType
$wb.ActiveChart.setSourceData($usedRange)
Try { $wb.ActiveChart.Name = "$($WorksheetName) - Chart" }
Catch { }
$wb.ActiveChart.Move([System.Reflection.Missing]::Value,$wb.Sheets.Item($ws.Name))
} Else {
$ws.Shapes.AddChart($ChartType).Chart.setSourceData($usedRange) | Out-Null
}
}
$wb.SaveAs($Path,$xlFixedFormat)
$wb.Close()
$xl.Quit()
While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)) {}
While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)) {}
If ($Title) { While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($TitleRange)) {} }
While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)) {}
While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)) {}
[GC]::Collect()
If ($PassThrough) { Return Get-Item $Path }
}
}
# End Export-Xlsx
####################
# Static Functions #
####################

# Clean Mac address
Function Clean-MacAddress
{
<#
	.SYNOPSIS
		Function to cleanup a MACAddress string
	
	.DESCRIPTION
		Function to cleanup a MACAddress string
	
	.PARAMETER MacAddress
		Specifies the MacAddress
	
	.PARAMETER Separator
		Specifies the separator every two characters
	
	.PARAMETER Uppercase
		Specifies the output must be Uppercase
	
	.PARAMETER Lowercase
		Specifies the output must be LowerCase
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:33:44:55'
	
		001122334455
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Uppercase
	
		001122DDEEFF
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Lowercase
	
		001122ddeeff
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Lowercase -Separator '-'
	
		00-11-22-dd-ee-ff
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Lowercase -Separator '.'
	
		00.11.22.dd.ee.ff
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Lowercase -Separator :
	
		00:11:22:dd:ee:ff
	
	.OUTPUTS
		System.String
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	[OutputType([String], ParameterSetName = "Upper")]
	[OutputType([String], ParameterSetName = "Lower")]
	[CmdletBinding(DefaultParameterSetName = 'Upper')]
	param
	(
		[Parameter(ParameterSetName = 'Lower')]
		[Parameter(ParameterSetName = 'Upper')]
		[String]$MacAddress,
		
		[Parameter(ParameterSetName = 'Lower')]
		[Parameter(ParameterSetName = 'Upper')]
		[ValidateSet(':', 'None', '.', "-")]
		$Separator,
		
		[Parameter(ParameterSetName = 'Upper')]
		[Switch]$Uppercase,
		
		[Parameter(ParameterSetName = 'Lower')]
		[Switch]$Lowercase
	)
	
	BEGIN
	{
		# Initial Cleanup
		$MacAddress = $MacAddress -replace "-", "" #Replace Dash
		$MacAddress = $MacAddress -replace ":", "" #Replace Colon
		$MacAddress = $MacAddress -replace "/s", "" #Remove whitespace
		$MacAddress = $MacAddress -replace " ", "" #Remove whitespace
		$MacAddress = $MacAddress -replace "\.", "" #Remove dots
		$MacAddress = $MacAddress.trim() #Remove space at the beginning
		$MacAddress = $MacAddress.trimend() #Remove space at the end
	}
	PROCESS
	{
		IF ($PSBoundParameters['Uppercase'])
		{
			$MacAddress = $macaddress.toupper()
		}
		IF ($PSBoundParameters['Lowercase'])
		{
			$MacAddress = $macaddress.tolower()
		}	
		IF ($PSBoundParameters['Separator'])
		{
			IF ($Separator -ne "None")
			{
				$MacAddress = $MacAddress -replace '(..(?!$))', "`$1$Separator"
			}
		}
	}
	END
	{
		Write-Output $MacAddress
	}
}
# End Clean Mac Address
# Get-IPAddress
Function Get-IPAddress
{
	Get-NetIPAddress | ?{($_.interfacealias -notlike "*loopback*") -and ($_.interfacealias -notlike "*vmware*") -and ($_.interfacealias -notlike "*loopback*") -and ($_.interfacealias -notlike "*bluetooth*") -and ($_.interfacealias -notlike "*isatap*")} | ft
}
# Reload Profile
Function Reload-Profile {
    @(
        $Profile.AllUsersAllHosts,
        $Profile.AllUsersCurrentHost,
        $Profile.CurrentUserAllHosts,
        $Profile.CurrentUserCurrentHost
    ) | % {
        if(Test-Path $_){
            Write-Verbose "Running $_"
            . $_
        }
    }    
}
# End Get-IPAddress


# DelProf2 User profiles
Function Remove-UserProfiles {

<#
.SYNOPSIS
    Written by: JBear 1/31/2017
	
    Remove user profiles from a specified system.

.DESCRIPTION
    Remove user profiles from a specified system with the use of DelProf2.exe.

.EXAMPLE
    Remove-UserProfiles Computer123456

        Note: Follow instructions and prompts to completetion.

#>

    param(
        [parameter(mandatory=$true)]
        [string[]]$computername
    )

 
    Function UseDelProf2 { 
               
        #Set parameters for remote computer and -WhatIf (/l)
        $WhatIf = @(

            "/l",
            "/c:$computer" 
        )
           
        #Runs DelProf2.exe with the /l parameter (or -WhatIf) to list potential User Profiles tagged for potential deletion
        & "C:\LazyWinAdmin\Win32-Tools\DelProf2.exe" $WhatIf

        #Display instructions on console
        Write-Host "`n`nPLEASE ENSURE YOU FULLY UNDERSTAND THIS COMMAND BEFORE USE `nTHIS WILL DELETE ALL USER PROFILE INFORMATION FOR SPECIFIED USER(S) ON THE SPECIFIED WORKSTATION!`n"

        #Prompt User for input
        $DeleteUsers = Read-Host -Prompt "To delete User Profiles, please use the following syntax ; Wildcards (*) are accepted. `nExample: /id:user1 /id:smith* /id:*john*`n `nEnter proper syntax to remove specific users" 

        #If only whitespace or a $null entry is entered, command is not run
        if([string]::IsNullOrWhiteSpace($DeleteUsers)) {

            Write-Host "`nImproper value entered, excluding all users from deletion. You will need to re-run the command on $computer, if you wish to try again...`n"

        }

        #If Read-Host contains proper syntax (Starts with /id:) run command to delete specified user; DelProf will give a confirmation prompt
        elseif($DeleteUsers -like "/id:*") {

            #Set parameters for remote computer
            $UserArgs = @(

                "/c:$computer"
            )

            #Split $DeleteUsers entries and add to $UserArgs array
            $UserArgs += $DeleteUsers.Split("")

            #Runs DelProf2.exe with $UserArgs parameters (i.e. & "C:\DelProf2.exe" /c:Computer1 /id:User1* /id:User7)
            & "C:\LazyWinAdmin\Win32-Tools\DelProf2.exe" $UserArgs
        }

        #If Read-Host doesn't begin with the input /id:, command is not run
        else {

            Write-Host "`nImproper value entered, excluding all users from deletion. You will need to re-run the command on $computer, if you wish to try again...`n"
        }
    }

    foreach($computer in $computername) {
        if(Test-Connection -Quiet -Count 1 -Computer $Computer) { 

            UseDelProf2 
        }

        else {
            
            Write-Host "`nUnable to connect to $computer. Please try again..." -ForegroundColor Red
        }

    }
}#End Remove-UserProfiles

# Remove-RemotePrintDrivers
Function Remove-RemotePrintDrivers {
  <# 
  .SYNOPSIS 
  Remove printer drivers from registry of specified workstation(s) 

  .EXAMPLE 
  Remove-RemotePrintDrivers Computer123456 

  .EXAMPLE 
  Remove-RemotePrintDrivers 123456 
  #> 
	param([Parameter(Mandatory=$true)]
	[string[]]$computername)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

Function RmPrintDrivers {

$i=0
$j=0
 	
foreach ($Computer in $ComputerName) { 

    Write-Progress -Activity "Clearing printer drivers..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

 
	Try {

		$RemoteSession = New-PSSession -ComputerName $Computer
}
	Catch {

		"Something went wrong. Unable to connect to $Computer"
		Break
}
	Invoke-Command -Session $RemoteSession -ScriptBlock {
    # Removes print drivers, other than default image drivers
		if ((Test-Path -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\') -eq $true) {
			Remove-Item -PATH 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3\*' -EXCLUDE "*ADOBE*", "*MICROSOFT*", "*XPS*", "*REMOTE*", "*FAX*", "*ONENOTE*" -recurse
			Remove-Item -PATH 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\*' -EXCLUDE "*ADOBE*", "*MICROSOFT*", "*XPS*", "*REMOTE*", "*FAX*", "*ONENOTE*" -recurse
		Set-Service Spooler -startuptype manual
		Restart-Service Spooler
		Set-Service Spooler -startuptype automatic
			}
		} -AsJob -JobName "ClearPrintDrivers"
	} 
} RmPrintDrivers | Wait-Job | Remove-Job

Remove-PSSession *

[Void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$RMprintConfirmation = [Microsoft.VisualBasic.Interaction]::MsgBox("Printer driver removal triggered on workstation(s)!", "OKOnly,SystemModal,Information", "Success")

}#End Remove-RemotePrintDrivers

# Remote Desktop Protocol
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
}# End RDP

# Get-LastReboot
Function Get-LastBoot {
  <# 
  .SYNOPSIS 
  Retrieve last restart time for specified workstation(s) 

  .EXAMPLE 
  Get-LastBoot Computer123456 

  .EXAMPLE 
  Get-LastBoot 123456 
  #> 
    param([Parameter(Mandatory=$true)]
	[string[]] $ComputerName)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

$i=0
$j=0

    foreach ($Computer in $ComputerName) {

    Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

 
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{
            "Computer Name" = $Computer
            "Last Reboot"= $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}# End Get-LastBoot


# Get-LoggedOnUser
Function Get-LoggedOnUser{
  <# 
  .SYNOPSIS 
  Retrieve current user logged into specified workstations(s) 

  .EXAMPLE 
  Get-LoggedOnUser Computer123456 

  .EXAMPLE 
  Get-LoggedOnUser 123456 
  #> 
	Param([Parameter(Mandatory=$true)]
	[string[]] $ComputerName)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}
	write-host("")
	write-host("Gathering resources. Please wait...")
	write-host("")

    $i=0
    $j=0

    foreach($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Logged On User..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

        $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer

        Write-Host "User Logged In: " $computerSystem.UserName "`n"
    }
}#End Get-LoggedOnUser

# Get-HotFixes
Function Get-HotFixes {
  <# 
  .SYNOPSIS 
  Grabs all processes on specified workstation(s).

  .EXAMPLE 
  Get-HotFixes Computer123456 

  .EXAMPLE 
  Get-HotFixes 123456 
  #> 
param (
    [Parameter(ValueFromPipeline=$true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [string]$NameRegex = '')

if(($computername.length -eq 6)) {
    [int32] $dummy_output = $null;

    if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        $computername = "Computer" + $computername.Replace("Computer","")
    }	
}

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

Function HotFix {

$i=0
$j=0

    foreach ($computer in $ComputerArray) {

        Write-Progress -Activity "Retrieving HotFix Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

        Get-HotFix -Computername $computer 
    }    
}

foreach ($computer in $ComputerName) {	     
    If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
        $ComputerArray += $Computer
    }	
}

$HotFix = HotFix
$DocPath = [environment]::getfolderpath("mydocuments") + "\HotFix-Report.csv"

    		Switch ($CheckBox.IsChecked){
    		    $true { $HotFix | Export-Csv $DocPath -NoTypeInformation -Force; }
    		    default { $HotFix | Out-GridView -Title "HotFix Report"; }
    		}

	if($CheckBox.IsChecked -eq $true){
	    Try { 
		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {
		 #Do Nothing 
	    }
	}
	
	else{
	    Try {
	        $listBox.Items.Add("$stamp HotFixes output processed!`n")
	    } 
	    Catch {
	        #Do Nothing 
	    }
	}
}# End Get-HotFixes

# Get-RemoteGroup Policies

Function Get-GPRemote {
  <# 
  .SYNOPSIS 
  Open Group Policy for specified workstation(s) 

  .EXAMPLE 
  Get-GPRemote Computer123456 

  .EXAMPLE 
  Get-GPRemote 123456 
  #> 
param(
[Parameter(Mandatory=$true)]
[string[]] $ComputerName)

if (($computername.length -eq 6)) {
    [int32] $dummy_output = $null;

    if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
       	$computername = "Computer" + $computername.Replace("Computer","")}	
}

$i=0
$j=0

foreach ($Computer in $ComputerName) {

    Write-Progress -Activity "Opening Remote Group Policy..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	#Opens (Remote) Group Policy for specified workstation
	gpedit.msc /gpcomputer: $Computer
    
	}
}#End Get-GPRemote

# Get Remote Processes

Function CheckProcess {
  <# 
  .SYNOPSIS 
  Grabs all processes on specified workstation(s).

  .EXAMPLE 
  CheckProcess Computer123456 

  .EXAMPLE 
  CheckProcess 123456 
  #> 
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = '')
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

Function ChkProcess {

$i=0
$j=0

    foreach ($computer in $ComputerArray) {

        Write-Progress -Activity "Retrieving System Processes..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

        $getProcess = Get-Process -ComputerName $computer

        foreach ($Process in $getProcess) {
                
             [pscustomobject]@{
		"Computer Name" = $computer
                "Process Name" = $Process.ProcessName
                PID = '{0:f0}' -f $Process.ID
                Company = $Process.Company
                "CPU(s)" = $Process.CPU
                Description = $Process.Description
             }           
         }
     } 
}
	
foreach ($computer in $ComputerName) {	     
    If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
        $ComputerArray += $Computer
    }	
}
	$chkProcess = ChkProcess | Sort "Computer Name" | Select "Computer Name","Process Name", PID, Company, "CPU(s)", Description
    	$DocPath = [environment]::getfolderpath("mydocuments") + "\Process-Report.csv"

    		Switch ($CheckBox.IsChecked){
    		    $true { $chkProcess | Export-Csv $DocPath -NoTypeInformation -Force; }
    		    default { $chkProcess | Out-GridView -Title "Processes";  }
    		}

	if($CheckBox.IsChecked -eq $true){
	    Try { 
		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {
		 #Do Nothing 
	    }
	}
	
	else{
	    Try {
	        $listBox.Items.Add("$stamp Check Process output processed!`n")
	    } 
	    Catch {
	        #Do Nothing 
	    }
	}
    
}#End CheckProcess

# Whois Check
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
} #end Function WhoIs

###########################
#
# netstat
# http://blogs.microsoft.co.il/blogs/scriptfanatic/archive/2011/02/10/How-to-find-running-processes-and-their-port-number.aspx
# Get-NetworkStatistics | where-object {$_.State -eq "LISTENING"} | Format-Table
###########################

Function Get-NetworkStatistics 
{ 
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
# End Get-NetworkStatistics

# Update SysInternals
<#
.Synopsis
   Download the latest sysinternals tools
.DESCRIPTION
   Downloads the latest sysinternals tools from https://live.sysinternals.com/ to a specified directory
   The Function downloads all .exe and .chm files available
.EXAMPLE
   Update-Sysinternals -Path C:\sysinternals
   Downloads the sysinternals tools to the directory C:\sysinternals
.EXAMPLE
   Update-Sysinternals -Path C:\Users\Matt\OneDrive\Tools\sysinternals
   Downloads the sysinternals tools to a user's OneDrive
#>
Function Update-Sysinternals {
    [CmdletBinding()]
    param (
        # Path to the directory were sysinternals tools will be downloaded to 
        [Parameter(Mandatory=$true)]      
        [string]
        $Path 
    )
    
    begin {
            if (-not (Test-Path -Path $Path)){
            Throw "The Path $_ does not exist"
        } else {
            $true
        }
        
            $uri = 'https://live.sysinternals.com/'
            $sysToolsPage = Invoke-WebRequest -Uri $uri
            
    }
    
    process {
        # create dir if it doesn't exist    
       
        Set-Location -Path $Path

        $sysTools = $sysToolsPage.Links.innerHTML | Where-Object -FilterScript {$_ -like "*.exe" -or $_ -like "*.chm"} 

        foreach ($sysTool in $sysTools){
            Invoke-WebRequest -Uri "$uri/$sysTool" -OutFile $sysTool
        }
    } #process
}
# End Update SysInternals

<#
#Function Get-UpTime
#{
#    param([string] $LastBootTime)
#    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
#    "$($Uptime.Days) days $($Uptime.Hours)h $($Uptime.Minutes)m"
#}
Function Get-Uptime {
# Accept input from the pipeline
Param([Parameter(mandatory=$true,ValueFromPipeline=$true)] [string[]]$ComputerName = @("."))

# Process the piped input (one computer at a time)
process { 

    # See if it responds to a ping, otherwise the WMI queries will fail
    $query = "select * from win32_pingstatus where address = '$ComputerName'"
    $ping = Get-WmiObject -query $query
if ($ping.protocoladdress) {
    # Ping responded, so connect to the computer via WMI
    $os = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ev myError -ea SilentlyContinue 

if ($myError -ne $null)
 {
  # Error: WMI did not respond
  "$ComputerName did not respond"
 } 
else
 { 
   $LastBootUpTime = $os.ConvertToDateTime($os.LastBootUpTime)
   $LocalDateTime = $os.ConvertToDateTime($os.LocalDateTime)
   
   # Calculate uptime - this is automatically a timespan
   $up = $LocalDateTime - $LastBootUpTime

   # Split into Days/Hours/Mins
   $uptime = "$($up.Days) days, $($up.Hours)h, $($up.Minutes)mins" 

   # Save the results for this computer in an object
   $results = new-object psobject
   $results | add-member noteproperty LastBootUpTime $LastBootUpTime
   $results | add-member noteproperty ComputerName $os.csname
   $results | add-member noteproperty uptime $uptime

   # Display the results
   $results | Select-Object ComputerName,LastBootUpTime, uptime
 }

# Next Ping result
}

# End of the process block
}}
#>

# End Get-Update
# Get AD GPO Replication
Function Get-ADGPOReplication
{
	<#
	.SYNOPSIS
		This Function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.DESCRIPTION
		This Function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.PARAMETER GPOName
		Specify the name of the GPO
	.PARAMETER All
		Specify that you want to retrieve all the GPO (slow if you have a lot of Domain Controllers)
	.EXAMPLE
		Get-ADGPOReplication -GPOName "Default Domain Policy"
	.EXAMPLE
		Get-ADGPOReplication -All
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		lazywinadmin.com
	
		VERSION HISTORY
		1.0 2014.09.22 	Initial version
						Adding some more Error Handling
						Fix some typo
	#>
	#requires -version 3

	[CmdletBinding()]
	PARAM (
		[parameter(Mandatory = $True, ParameterSetName = "One")]
		[String[]]$GPOName,
		[parameter(Mandatory = $True, ParameterSetName = "All")]
		[Switch]$All
	)
Remove-Module Carbon
	BEGIN
	{
		TRY
		{
			if (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction Stop -ErrorVariable ErrorBeginIpmoAD }
			if (-not (Get-Module -Name GroupPolicy)) { Import-Module -Name GroupPolicy -ErrorAction Stop -ErrorVariable ErrorBeginIpmoGP }
		}
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			IF ($ErrorBeginIpmoAD) { Write-Warning -Message "[BEGIN] Error while Importing the module Active Directory" }
			IF ($ErrorBeginIpmoGP) { Write-Warning -Message "[BEGIN] Error while Importing the module Group Policy" }
			Write-Warning -Message "[BEGIN] $($Error[0].exception.message)"
		}
	}
	PROCESS
	{
		FOREACH ($DomainController in ((Get-ADDomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetDC -filter *).hostname))
		{
			TRY
			{
				IF ($psBoundParameters['GPOName'])
				{
					Foreach ($GPOItem in $GPOName)
					{
						$GPO = Get-GPO -Name $GPOItem -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPO
						
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPOItem
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}#Foreach ($GPOItem in $GPOName)
				}#IF ($psBoundParameters['GPOName'])
				IF ($psBoundParameters['All'])
				{
					$GPOList = Get-GPO -All -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPOAll
					
					foreach ($GPO in $GPOList)
					{
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPO.DisplayName
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}
				}#IF ($psBoundParameters['All'])
			}#TRY
			CATCH
			{
				Write-Warning -Message "[PROCESS] Something wrong happened"
				IF ($ErrorProcessGetDC) { Write-Warning -Message "[PROCESS] Error while running retrieving Domain Controllers with Get-ADDomainController" }
				IF ($ErrorProcessGetGPO) { Write-Warning -Message "[PROCESS] Error while running Get-GPO" }
				IF ($ErrorProcessGetGPOAll) { Write-Warning -Message "[PROCESS] Error while running Get-GPO -All" }
				Write-Warning -Message "[PROCESS] $($Error[0].exception.message)"
			}
		}#FOREACH
	}#PROCESS
}
# End Get-ADGPOReplication
# Test Registry Value
Function Test-RegistryValue {

param (

 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Path,

[parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Value
)

try {

Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
 return $true
 }

catch {

return $false

}

}
### End Test-RegistryValue
# Get Local Admins 
Function Get-LocalAdmin { 
param ($ComputerName) 
 
$admins = Gwmi win32_groupuser –computer $ComputerName  
$admins = $admins |? {$_.groupcomponent –like '*"Administrators"'} 
 
$admins |% { 
$_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul 
$matches[1].trim('"') + “\” + $matches[2].trim('"') 
} 
}
### End Get-LocalAdmin
####################
# Static Functions #
####################
# Touch
Function touch { $args | foreach-object {write-host > $_} }
# Notepad++
Function NPP { Start-Process -FilePath "${Env:ProgramFiles(x86)}\Notepad++\Notepad++.exe" }#-ArgumentList $args }
# Find File
Function findfile($name) {
	ls -recurse -filter "*${name}*" -ErrorAction SilentlyContinue | foreach {
		$place_path = $_.directory
		echo "${place_path}\${_}"
	}
}
# RM -RF
Function rm-rf($item) { Remove-Item $item -Recurse -Force }
# SUDO
Function sudo(){
	Invoke-Elevated @args
}
# SED
Function PSsed($file, $find, $replace){
	(Get-Content $file).replace("$find", $replace) | Set-Content $file
}
# SED-Recursive
Function PSsed-recursive($filePattern, $find, $replace) {
	$files = ls . "$filePattern" -rec # -Exclude
	foreach ($file in $files) {
		(Get-Content $file.PSPath) |
		Foreach-Object { $_ -replace "$find", "$replace" } |
		Set-Content $file.PSPath
	}
}
# PSGrep
Function PSgrep {

    [CmdletBinding()]
    Param(
    
        # source file to grep
        [Parameter(Mandatory=$true)]
        [string]$SourceFileName, 

        # string to search for
        [Parameter(Mandatory=$true)]
        [string]$SearchStrings,

        # do we write to file
        [Parameter()]
        [string]$OutputFile
    )

        # break the comma separated strings up
        $Strings = @()
        $Strings = $SearchStrings.split(',')
        $count = 0

        # write-host $Strings

        $Content = Get-Content $SourceFileName
        
        $Content | ForEach-Object { 
            foreach ($String in $Strings) {
                # $String
                if($_ -match $String){
                    $count ++
                    if (!($OutputFile)) {
                        write-host $_
                    } else {
                        $_ | Out-File -FilePath ".\$($OutputFile)" -Append -Force
                }

            }

        }

    }

    Write-Host "$($Count) matches found"
}
# End PSgrep
# Which
Function which($name) {
	Get-Command $name | Select-Object -ExpandProperty Definition
}
# Cut
Function cut(){
	foreach ($part in $input) {
		$line = $part.ToString();
		$MaxLength = [System.Math]::Min(200, $line.Length)
		$line.subString(0, $MaxLength)
	}
}
# Search Text Files
Function Search-AllTextFiles {
    param(
        [parameter(Mandatory=$true,position=0)]$Pattern, 
        [switch]$CaseSensitive,
        [switch]$SimpleMatch
    );

    Get-ChildItem . * -Recurse -Exclude ('*.dll','*.pdf','*.pdb','*.zip','*.exe','*.jpg','*.gif','*.png','*.ico','*.svg','*.bmp','*.psd','*.cache','*.doc','*.docx','*.xls','*.xlsx','*.dat','*.mdf','*.nupkg','*.snk','*.ttf','*.eot','*.woff','*.tdf','*.gen','*.cfs','*.map','*.min.js','*.data') | Select-String -Pattern:$pattern -SimpleMatch:$SimpleMatch -CaseSensitive:$CaseSensitive
}
# Add to Zip
Function AddTo-7zip($zipFileName) {
    BEGIN {
        #$7zip = "$($env:ProgramFiles)\7-zip\7z.exe"
        $7zip = Find-Program "\7-zip\7z.exe"
		if(!([System.IO.File]::Exists($7zip))){
			throw "7zip not found";
		}
    }
    PROCESS {
        & $7zip a -tzip $zipFileName $_
    }
    END {
    }
}
## End Add to Zip

# Connect to Exchange
Function GoGo-PSExch {

    param(
        [Parameter( Mandatory=$false)]
        [string]$URL="USONVSVREX01"
    )
    
    #$Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

    $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos #-Credential #$Credentials

    Import-PSSession $ExOPSession -AllowClobber

}
## End Connect to Exchange
## Out-File in UTF8 NonBom
Function Out-FileUtf8NoBom {

  [CmdletBinding()]
  param(
    [Parameter(Mandatory, Position=0)] [string] $LiteralPath,
    [switch] $Append,
    [switch] $NoClobber,
    [AllowNull()] [int] $Width,
    [Parameter(ValueFromPipeline)] $InputObject
  )

  #requires -version 3

  # Make sure that the .NET framework sees the same working dir. as PS
  # and resolve the input path to a full path.
  [System.IO.Directory]::SetCurrentDirectory($PWD) # Caveat: .NET Core doesn't support [Environment]::CurrentDirectory
  $LiteralPath = [IO.Path]::GetFullPath($LiteralPath)

  # If -NoClobber was specified, throw an exception if the target file already
  # exists.
  if ($NoClobber -and (Test-Path $LiteralPath)) {
    Throw [IO.IOException] "The file '$LiteralPath' already exists."
  }

  # Create a StreamWriter object.
  # Note that we take advantage of the fact that the StreamWriter class by default:
  # - uses UTF-8 encoding
  # - without a BOM.
  $sw = New-Object IO.StreamWriter $LiteralPath, $Append

  $htOutStringArgs = @{}
  if ($Width) {
    $htOutStringArgs += @{ Width = $Width }
  }

  # Note: By not using begin / process / end blocks, we're effectively running
  #       in the end block, which means that all pipeline input has already
  #       been collected in automatic variable $Input.
  #       We must use this approach, because using | Out-String individually
  #       in each iteration of a process block would format each input object
  #       with an indvidual header.
  try {
    $Input | Out-String -Stream @htOutStringArgs | % { $sw.WriteLine($_) }
  } finally {
    $sw.Dispose()
  }

}
## End Out-File in UTF8 NonBom

## Invoke VBScript
Function Invoke-VBScript {
    <#
    .Synopsis
       Run VBScript from PowerShell
    .DESCRIPTION
       Used to invoke VBScript from PowerShell

       Will run the VBScript in a separate job using cscript.exe
    .PARAMETER Path
       Path to VBScript.
       Accepts relative or absolute path.
    .PARAMETER Argument
       Arguments to pass to VBScript
    .PARAMETER Wait
       Wait for VBScript to finish   
    .EXAMPLE
       Invoke-VBScript -Path '.\VBScript1.vbs' -Arguments '"MyFirstArgument"', '"MySecondArgument"' -Wait
       Run VBScript1.vbs using cscript and wait for the script to complete.
       Displays progressbar while waiting.
       Returns script output as single string.
    .EXAMPLE
       '.\VBScript1.vbs', '.\VBScript2.vbs' | Invoke-VBScript -Arguments '"MyArgument"'
       Starts both VBScript1.vbs and VBScript2.vbs in separate jobs simultaneously.
       Both scripts will be run using the same arguments.
       Returns job items.
    .EXAMPLE
       [PSCustomObject]@{Path='.\VBScript1.vbs';Arguments='"Script1"'},[PSCustomObject]@{Path='.\VBScript2.vbs';Arguments='"Script2"'} | Invoke-VBScript -Wait -Verbose
       Runs two scripts after each other, waiting to one to complete
       before starting next.
       Each script will run with different parameters.
       Displays progressbar while waiting.
       Returns script output in one single string per script.
    .NOTES
       Written by Simon Wåhlin
       http://blog.simonw.se
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='None',PositionalBinding=$false)]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [ValidateScript({if(Test-Path $_){$true}else{Throw "Could not find script: [$_]"}})]
        [String]
        $Path,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Args')]
        [String[]]
        $Argument,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch]
        $Wait
    )
    Begin
    {
        Write-Verbose -Message 'Locating cscript.exe'
        $cscriptpath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\cscript.exe'
        if(-Not(Test-Path -Path $cscriptpath))
        {
            Throw 'cscript.exe not found.'
        }
        Write-Verbose -Message ('cscript.exe found in: {0}' -f $cscriptpath)
    }
    Process
    {
        Try
        {
            $ResolvedPath = Resolve-Path -Path $Path
            Write-Verbose -Message ('Processing script: {0}' -f $ResolvedPath)
            if($PSBoundParameters.ContainsKey('Argument'))
            {
                $ScriptBlock = [scriptblock]::Create(('& "{0}" "{1}" "{2}"' -f $cscriptpath, $ResolvedPath,($Argument -join '" "')))
            }
            else
            {
                $ScriptBlock = $ScriptBlock = [scriptblock]::Create(('& "{0}" "{1}"' -f $cscriptpath, $ResolvedPath))
            }
            Write-Verbose -Message 'Starting script'
            if($PSCmdlet.ShouldProcess($ResolvedPath,'Invoke script'))
            {
                $Job = Start-Job -ScriptBlock $ScriptBlock
                if($Wait)
                {
                    $Activity = 'Waiting for script to complete: {0}' -f $ResolvedPath
                    Write-Progress -Activity $Activity -Id 1
                    $i = 1
                    While($Job.State -eq 'Running')
                    {
                        $WaitTime = (Get-Date) - $Job.PSBeginTime
                        Write-Progress -Activity $Activity -Status "Waited for $($WaitTime.TotalSeconds -as [int]) seconds." -Id 1 -PercentComplete ($i%100)
                        Start-Sleep -Seconds 1
                        $i++
                    }
                    Write-Progress -Activity $Activity -Status 'Waiting' -Id 1 -Completed
                    $Result = Foreach($JobInstance in ($Job,$Job.ChildJobs))
                    {
                        if($JobInstance.Error -ne $null)
                        {
                            Throw $JobInstance.Error.Exception.Message
                        }
                        else
                        {
                            $JobInstance.Output
                        }
                    }
                    Write-Output -InputObject ($Result -join "`n")
                    Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
                }
                else
                {
                    Write-Output -InputObject $Job
                }
            }
            Write-Verbose -Message 'Finished processing script'
        }
        Catch
        {
            Throw
        }
    }
}
## End Invoke VBScript
## Function Get-MOTD
Function Get-MOTD {

<#
.NAME
    Get-MOTD
.SYNOPSIS
    Displays system information to a host.
.DESCRIPTION
    The Get-MOTD cmdlet is a system information tool written in PowerShell. 
.EXAMPLE
#>


  [CmdletBinding()]
	
  Param(
    [Parameter(Position=0,Mandatory=$false)]
	[ValidateNotNullOrEmpty()]
    [string[]]$ComputerName
    ,
    [Parameter(Position=1,Mandatory=$false)]
    [PSCredential]
    [System.Management.Automation.CredentialAttribute()]$Credential
  )

  Begin {
	
        If (-Not $ComputerName) {
            $RemoteSession = $null
        }
        #Define ScriptBlock for data collection
        $ScriptBlock = {
            $Operating_System = Get-CimInstance -ClassName Win32_OperatingSystem
            $Logical_Disk = Get-CimInstance -ClassName Win32_LogicalDisk |
            Where-Object -Property DeviceID -eq $Operating_System.SystemDrive
			Try {
				$PCLi = Get-PowerCLIVersion
				$PCLiVer = ' | PowerCLi ' + [string]$PCLi.Major + '.' + [string]$PCLi.Minor + '.' + [string]$PCLi.Revision + '.' + [string]$PCLi.Build
			} Catch {$PCLiVer = ''}
			If ($DomainName = ([System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()).DomainName) {$DomainName = '.' + $DomainName}
			
            [pscustomobject]@{
                Operating_System = $Operating_System
                Processor = Get-CimInstance -ClassName Win32_Processor
                Process_Count = (Get-Process).Count
                Shell_Info = ("{0}.{1}" -f $PSVersionTable.PSVersion.Major,$PSVersionTable.PSVersion.Minor) + $PCLiVer
                Logical_Disk = $Logical_Disk
            }
        }
  } #End Begin

  Process {
	
        If ($ComputerName) {
            If ("$ComputerName" -ne "$env:ComputerName") {
                # Build Hash to be used for passing parameters to 
                # New-PSSession commandlet
                $PSSessionParams = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }

                # Add optional parameters to hash
                If ($Credential) {
                    $PSSessionParams.Add('Credential', $Credential)
                }

                # Create remote powershell session   
                Try {
                    $RemoteSession = New-PSSession @PSSessionParams
                }
                Catch {
                    Throw $_.Exception.Message
                }
            } Else { 
                $RemoteSession = $null
            }
        }
        
        # Build Hash to be used for passing parameters to 
        # Invoke-Command commandlet
        $CommandParams = @{
            ScriptBlock = $ScriptBlock
            ErrorAction = 'Stop'
        }
        
        # Add optional parameters to hash
        If ($RemoteSession) {
            $CommandParams.Add('Session', $RemoteSession)
        }
               
        # Run ScriptBlock    
        Try {
            $ReturnedValues = Invoke-Command @CommandParams
        }
        Catch {
            If ($RemoteSession) {
            	Remove-PSSession $RemoteSession
            }
            Throw $_.Exception.Message
        }

        # Assign variables
        Import-Module MS-Module
        $Date = Get-Date
        $OS_Name = $ReturnedValues.Operating_System.Caption + ' [Installed: ' + ([datetime]$ReturnedValues.Operating_System.InstallDate).ToString('dd-MMM-yyyy') + ']'
        $Computer_Name = $ReturnedValues.Operating_System.CSName
		If ($DomainName) {$Computer_Name = $Computer_Name + $DomainName.ToUpper()}
        $Kernel_Info = $ReturnedValues.Operating_System.Version + ' [' + $ReturnedValues.Operating_System.OSArchitecture + ']'
        $Process_Count = $ReturnedValues.Process_Count
        $Uptime = "$(($Uptime = $Date - $($ReturnedValues.Operating_System.LastBootUpTime)).Days) days, $($Uptime.Hours) hours, $($Uptime.Minutes) minutes"
        $Shell_Info = $ReturnedValues.Shell_Info
        $CPU_Info = $ReturnedValues.Processor.Name -replace '\(C\)', '' -replace '\(R\)', '' -replace '\(TM\)', '' -replace 'CPU', '' -replace '\s+', ' '
        $Current_Load = $ReturnedValues.Processor.LoadPercentage    
        $Memory_Size = "{0} MB/{1} MB " -f (([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))-
        ([math]::round($ReturnedValues.Operating_System.FreePhysicalMemory/1KB))),([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))
		$Disk_Size = "{0} GB/{1} GB" -f (([math]::round($ReturnedValues.Logical_Disk.Size/1GB)-
        [math]::round($ReturnedValues.Logical_Disk.FreeSpace/1GB))),([math]::round($ReturnedValues.Logical_Disk.Size/1GB))

        # Write to the Console
        Write-Host -Object ("")
        Write-Host -Object ("")
        Write-Host -Object ("         ,.=:^!^!t3Z3z.,                  ") -ForegroundColor Red
        Write-Host -Object ("        :tt:::tt333EE3                    ") -ForegroundColor Red
        Write-Host -Object ("        Et:::ztt33EEE ") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @Ee.,      ..,     $($Date.ToString('dd-MMM-yyyy HH:mm:ss'))") -ForegroundColor Green
        Write-Host -Object ("       ;tt:::tt333EE7") -NoNewline -ForegroundColor Red
        Write-Host -Object (" ;EEEEEEttttt33#     ") -ForegroundColor Green
        Write-Host -Object ("      :Et:::zt333EEQ.") -NoNewline -ForegroundColor Red
        Write-Host -Object (" SEEEEEttttt33QL     ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("User: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$env:USERDOMAIN\$env:UserName") -ForegroundColor Cyan
        Write-Host -Object ("      it::::tt333EEF") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @EEEEEEttttt33F      ") -NoNewline -ForeGroundColor Green
        Write-Host -Object ("Hostname: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Computer_Name") -ForegroundColor Cyan
        Write-Host -Object ("     ;3=*^``````'*4EEV") -NoNewline -ForegroundColor Red
        Write-Host -Object (" :EEEEEEttttt33@.      ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("OS: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$OS_Name") -ForegroundColor Cyan
        Write-Host -Object ("     ,.=::::it=., ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("``") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @EEEEEEtttz33QF       ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("Kernel: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("NT ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("$Kernel_Info") -ForegroundColor Cyan
        Write-Host -Object ("    ;::::::::zt33) ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("  '4EEEtttji3P*        ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("Uptime: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Uptime") -ForegroundColor Cyan
        Write-Host -Object ("   :t::::::::tt33.") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (":Z3z.. ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object (" ````") -NoNewline -ForegroundColor Green
        Write-Host -Object (" ,..g.        ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Shell: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("PowerShell $Shell_Info") -ForegroundColor Cyan
        Write-Host -Object ("   i::::::::zt33F") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" AEEEtttt::::ztF         ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("CPU: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$CPU_Info") -ForegroundColor Cyan
        Write-Host -Object ("  ;:::::::::t33V") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" ;EEEttttt::::t3          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Processes: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Process_Count") -ForegroundColor Cyan
        Write-Host -Object ("  E::::::::zt33L") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" @EEEtttt::::z3F          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Current Load: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Current_Load") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("%") -ForegroundColor Cyan
        Write-Host -Object (" {3=*^``````'*4E3)") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" ;EEEtttt:::::tZ``          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Memory: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Memory_Size`t") -ForegroundColor Cyan -NoNewline
		New-PercentageBar -DrawBar -Value (([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))-([math]::round($ReturnedValues.Operating_System.FreePhysicalMemory/1KB))) -MaxValue ([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB)); "`r"
        Write-Host -Object ("             ``") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" :EEEEtttt::::z7            ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("System Volume: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Disk_Size`t") -ForegroundColor Cyan -NoNewline
		New-PercentageBar -DrawBar -Value (([math]::round($ReturnedValues.Logical_Disk.Size/1GB)-[math]::round($ReturnedValues.Logical_Disk.FreeSpace/1GB))) -MaxValue ([math]::round($ReturnedValues.Logical_Disk.Size/1GB)); "`r"
        Write-Host -Object ("                 'VEzjt:;;z>*``           ") -ForegroundColor Yellow
        Write-Host -Object ("                      ````                  ") -ForegroundColor Yellow
        Write-Host -Object ("")
  } #End Process

  End {
        If ($RemoteSession) {
            Remove-PSSession $RemoteSession
        }
  }
} #End Function Get-MOTD

## Change Attributes
Function Get-FileAttribute{
    param($file,$attribute)
    $val = [System.IO.FileAttributes]$attribute;
    if((gci $file -force).Attributes -band $val -eq $val){$true;} else { $false; }
} 


Function Set-FileAttribute{
    param($file,$attribute)
    $file =(gci $file -force);
    $file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]$attribute).value__;
    if($?){$true;} else {$false;}
} 

## End Change Attributes

## Remote Group Policy
Function GPR {
<# 
.SYNOPSIS 
    Open Group Policy for specified workstation(s) 

.EXAMPLE 
    GPR Computer123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [String[]]$ComputerName,

    $i=0,
    $j=0
)

    foreach ($Computer in $ComputerName) {

        Write-Progress -Activity "Opening Remote Group Policy..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

        #Opens (Remote) Group Policy for specified workstation
        GPedit.msc /gpcomputer: $Computer
    }
}#End GPR

## Begin Lastboot

Function LastBoot {
<# 
.SYNOPSIS 
    Retrieve last restart time for specified workstation(s) 

.EXAMPLE 
    LastBoot Computer123456 

.EXAMPLE 
    LastBoot 123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [String[]]$ComputerName,

    $i=0,
    $j=0
)

    foreach($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)
 
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{

            ComputerName = $Computer
            LastReboot = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}#End LastBoot

#Begin SYSinfo
Function SYSinfo {
<# 
.SYNOPSIS 
  Retrieve basic system information for specified workstation(s) 

.EXAMPLE 
  SYS Computer123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [string[]] $ComputerName,
    
    $i=0,
    $j=0
)

$Stamp = (Get-Date -Format G) + ":"

    Function Systeminformation {
	
        foreach ($Computer in $ComputerName) {

            if(!([String]::IsNullOrWhiteSpace($Computer))) {

                if(Test-Connection -Quiet -Count 1 -Computer $Computer) {

                    Write-Progress -Activity "Getting Sytem Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	                Start-Job -ScriptBlock { param($Computer) 

	                    #Gather specified workstation information; CimInstance only works on 64-bit
	                    $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer
	                    $computerBIOS = Get-CimInstance CIM_BIOSElement -Computer $Computer
	                    $computerOS = Get-CimInstance CIM_OperatingSystem -Computer $Computer
	                    $computerCPU = Get-CimInstance CIM_Processor -Computer $Computer
	                    $computerHDD = Get-CimInstance Win32_LogicalDisk -Computer $Computer -Filter "DeviceID = 'C:'"
    
                        [PSCustomObject] @{

                            ComputerName = $computerSystem.Name
                            LastReboot = $computerOS.LastBootUpTime
                            OperatingSystem = $computerOS.OSArchitecture + " " + $computerOS.caption
                            Model = $computerSystem.Model
                            RAM = "{0:N2}" -f [int]($computerSystem.TotalPhysicalMemory/1GB) + "GB"
                            DiskCapacity = "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
                            TotalDiskSpace = "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
                            CurrentUser = $computerSystem.UserName
                        }
                    } -ArgumentList $Computer
                }

                else {

                    Start-Job -ScriptBlock { param($Computer)  
                     
                        [PSCustomObject] @{

                            ComputerName=$Computer
                            LastReboot="Unable to PING."
                            OperatingSystem="$Null"
                            Model="$Null"
                            RAM="$Null"
                            DiskCapacity="$Null"
                            TotalDiskSpace="$Null"
                            CurrentUser="$Null"
                        }
                    } -ArgumentList $Computer                       
                }
            }

            else {
                 
                Start-Job -ScriptBlock { param($Computer)  
                     
                    [PSCustomObject] @{

                        ComputerName = "Value is null."
                        LastReboot = "$Null"
                        OperatingSystem = "$Null"
                        Model = "$Null"
                        RAM = "$Null"
                        DiskCapacity = "$Null"
                        TotalDiskSpace = "$Null"
                        CurrentUser = "$Null"
                    }
                } -ArgumentList $Computer
            }
        } 
    }

    $SystemInformation = SystemInformation | Receive-Job -Wait | Select ComputerName, CurrentUser, OperatingSystem, Model, RAM, DiskCapacity, TotalDiskSpace, LastReboot
    $DocPath = [environment]::getfolderpath("mydocuments") + "\SystemInformation-Report.csv"

	Switch($CheckBox.IsChecked) {

		$true { 
            
            $SystemInformation | Export-Csv $DocPath -NoTypeInformation -Force 
        }

		default { 
            
            $SystemInformation | Out-GridView -Title "System Information"
        }
    }

	if($CheckBox.IsChecked -eq $true) {

	    Try { 

		    $listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {

		    #Do Nothing 
	    }
	}
	
	else {

	    Try {

	        $listBox.Items.Add("$stamp System Information output processed!`n")
	    } 

	    Catch {

	        #Do Nothing 
	    }
	}
}#End SYSinfo

#Begin NetMessage

Function NetMSG {
<# 
.SYNOPSIS 
    Generate a pop-up window on specified workstation(s) with desired message 

.EXAMPLE 
    NetMSG Computer123456 
#> 
	
param(

    [Parameter(Mandatory=$true)]
    [String[]] $ComputerName,

    [Parameter(Mandatory=$true,HelpMessage='Enter desired message')]
    [String]$MyMessage,

    [String]$User = [Environment]::UserName,

    [String]$UserJob = (Get-ADUser $User -Property Title).Title,
    
    [String]$CallBack = "$User | 5-2444 | $UserJob",

    $i=0,
    $j=0
)

    Function SendMessage {

        foreach($Computer in $ComputerName) {

            Write-Progress -Activity "Sending messages..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)         

            #Invoke local MSG command on specified workstation - will generate pop-up message for any user logged onto that workstation - *Also shows on Login screen, stays there for 100,000 seconds or until interacted with
            Invoke-Command -ComputerName $Computer { param($MyMessage, $CallBack, $User, $UserJob)
 
                MSG /time:100000 * /v "$MyMessage {$CallBack}"
            } -ArgumentList $MyMessage, $CallBack, $User, $UserJob -AsJob
        }
    }

    SendMessage | Wait-Job | Remove-Job

}#End NetMSG

Function InstallApplication {

<#     
.SYNOPSIS     
  
    Copies and installs specifed filepath ($Path). This serves as a template for the following filetypes: .EXE, .MSI, & .MSP 

.DESCRIPTION     
    Copies and installs specifed filepath ($Path). This serves as a template for the following filetypes: .EXE, .MSI, & .MSP

.EXAMPLE    
    .\InstallAsJob (Get-Content C:\ComputerList.txt)

.EXAMPLE    
    .\InstallAsJob Computer1, Computer2, Computer3 
    
.NOTES   
    Author: JBear 
    Date: 2/9/2017 
    
    Edit: JBear
    Date: 10/13/2017 
#> 

param(

    [Parameter(Mandatory=$true,HelpMessage="Enter Computername(s)")]
    [String[]]$Computername,

    [Parameter(ValueFromPipeline=$true,HelpMessage="Enter installer path(s)")]
    [String[]]$Path = $null,

    [Parameter(ValueFromPipeline=$true,HelpMessage='Enter remote destination: C$\Directory')]
    $Destination = "C$\TempApplications"
)

    if($Path -eq $null) {

        Add-Type -AssemblyName System.Windows.Forms

        $Dialog = New-Object System.Windows.Forms.OpenFileDialog
        $Dialog.InitialDirectory = "\\lasfs03\Software\Current Version\Deploy"
        $Dialog.Title = "Select Installation File(s)"
        $Dialog.Filter = "Installation Files (*.exe,*.msi,*.msp)|*.exe; *.msi; *.msp"        
        $Dialog.Multiselect=$true
        $Result = $Dialog.ShowDialog()

        if($Result -eq 'OK') {

            Try {
        
                $Path = $Dialog.FileNames
            }

            Catch {

                $Path = $null
	            Break
            }
        }

        else {

            #Shows upon cancellation of Save Menu
            Write-Host -ForegroundColor Yellow "Notice: No file(s) selected."
            Break
        }
    }

    #Create Function    
    Function InstallAsJob {

        #Each item in $Computernam variable        
        foreach($Computer in $Computername) {

            #If $Computer IS NOT null or only whitespace
            if(!([string]::IsNullOrWhiteSpace($Computer))) {

                #Test-Connection to $Computer
                if(Test-Connection -Quiet -Count 1 $Computer) {                                               
                     
                    #Create job on localhost
                    Start-Job { param($Computer, $Path, $Destination)

                        foreach($P in $Path) {
                            
                            #Static Temp location
                            $TempDir = "\\$Computer\$Destination"

                            #Create $TempDir directory
                            if(!(Test-Path $TempDir)) {

                                New-Item -Type Directory $TempDir | Out-Null
                            }
                     
                            #Retrieve Leaf object from $Path
                            $FileName = (Split-Path -Path $P -Leaf)

                            #New Executable Path
                            $Executable = "C:\$(Split-Path -Path $Destination -Leaf)\$FileName"

                            #Copy needed installer files to remote machine
                            Copy-Item -Path $P -Destination $TempDir

                            #Install .EXE
                            if($FileName -like "*.exe") {

                                Function InvokeEXE {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
                                    
                                        Try {

                                            #Start EXE file
                                            Start-Process $Executable -ArgumentList "/s" -Wait -NoNewWindow
                                            
                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }
                                       
                                    } -AsJob -JobName "Silent EXE Install" -ArgumentList $TempDir, $FileName, $Executable
                                }

                                InvokeEXE | Receive-Job -Wait
                            }
                               
                            #Install .MSI                                        
                            elseif($FileName -like "*.msi") {

                                Function InvokeMSI {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
				    
				    	                $MSIArguments = @(
						
						                    "/i"
						                    $Executable
						                    "/qn"
					                    )

                                        Try {
                                        
                                            #Start MSI file                                    
                                            Start-Process 'msiexec.exe' -ArgumentList $MSIArguments -Wait -ErrorAction Stop

                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }                              
                                    } -AsJob -JobName "Silent MSI Install" -ArgumentList $TempDir, $FileName, $Executable                            
                                }

                                InvokeMSI | Receive-Job -Wait
                            }

                            #Install .MSP
                            elseif($FileName -like "*.msp") { 
                                                                       
                                Function InvokeMSP {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
				    
				    	                $MSPArguments = @(
						
						                    "/p"
						                    $Executable
						                    "/qn"
					                    )				    

                                        Try {
                                                                                
                                            #Start MSP file                                    
                                            Start-Process 'msiexec.exe' -ArgumentList $MSPArguments -Wait -ErrorAction Stop

                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }                             
                                    } -AsJob -JobName "Silent MSP Installer" -ArgumentList $TempDir, $FileName, $Executable
                                }

                                InvokeMSP | Receive-Job -Wait
                            }

                            else {

                                Write-Host "$Destination has an unsupported file extension. Please try again."                        
                            }
                        }                      
                    } -Name "Application Install" -Argumentlist $Computer, $Path, $Destination            
                }
                                            
                else {                                
                    
                    Write-Host "Unable to connect to $Computer."                
                }            
            }        
        }   
    }

    #Call main Function
    InstallAsJob
    Write-Host "`nJob creation complete. Please use the Get-Job cmdlet to check progress.`n"
    Write-Host "Once all jobs are complete, use Get-Job | Receive-Job to retrieve any output or, Get-Job | Remove-Job to clear jobs from the session cache."
}#End InstallApplication

# Begin Get-Icon

Function Get-Icon {
    <#
        .SYNOPSIS
            Gets the icon from a file

        .DESCRIPTION
            Gets the icon from a file and displays it in a variety formats.

        .PARAMETER Path
            The path to a file to get the icon

        .PARAMETER ToBytes
            Displays outputs as a byte array

        .PARAMETER ToBitmap
            Display the icon as a bitmap object

        .PARAMETER ToBase64
            Displays the icon in Base64 encoded format

        .NOTES
            Name: Get-Icon
            Author: Boe Prox
            Version History:
                1.0 //Boe Prox - 11JAN2016
                    - Initial version

        .OUTPUT
            System.Drawing.Icon
            System.Drawing.Bitmap
            System.String
            System.Byte[]

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe'

            FullName : C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe
            Handle   : 164169893
            Height   : 32
            Size     : {Width=32, Height=32}
            Width    : 32

            Description
            -----------
            Returns the System.Drawing.Icon representation of the icon

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBitmap

            Tag                  : 
            PhysicalDimension    : {Width=32, Height=32}
            Size                 : {Width=32, Height=32}
            Width                : 32
            Height               : 32
            HorizontalResolution : 96
            VerticalResolution   : 96
            Flags                : 2
            RawFormat            : [ImageFormat: b96b3caa-0728-11d3-9d7b-0000f81ef32e]
            PixelFormat          : Format32bppArgb
            Palette              : System.Drawing.Imaging.ColorPalette
            FrameDimensionsList  : {7462dc86-6180-4c7e-8e3f-ee7333a7a483}
            PropertyIdList       : {}
            PropertyItems        : {}

            Description
            -----------
            Returns the System.Drawing.Bitmap representation of the icon

        .EXAMPLE
            $FileName = 'C:\Temp\PowerShellIcon.png'
            $Format = [System.Drawing.Imaging.ImageFormat]::Png
            (Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBitmap).Save($FileName,$Format)

            Description
            -----------
            Saves the icon as a file.

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBase64

            AAABAAEAICAQHQAAAADoAgAAFgAAACgAAAAgAAAAQAAAAAEABAAAAAAAgAIAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAACAAACAAAAAgIAAgAAAAIAAgACAgAAAgICAAMDAwAAAAP8AAP8AAAD//wD/AAAA/wD/AP
            //AAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABmZmZmZmZmZmZmZgAAAAAAaId3d3d3d4iIiIdgAA
            AHdmhmZmZmZmZmZmZoZAAAB2ZnZmZmZmZmZmZmZ3YAAAdmZ3ZmiHZniIiHZmaGAAAHZmd2Zv/4eIiIi
            GZmhgAAB2ZmdmZ4/4eIh3ZmZnYAAAd2ZnZmZo//h2ZmZmZ3YAAHZmaGZmZo//h2ZmZmd2AAB3Zmd2Zm
            Znj/h2ZmZmhgAAd3dndmZmZuj/+GZmZoYAAHd3dod3dmZuj/9mZmZ2AACHd3aHd3eIiP/4ZmZmd2AAi
            Hd2iIiIiI//iId2ZndgAIiIhoiIiIj//4iIiIiIYACIiId4iIiP//iIiIiIiGAAiIiIaIiI//+IiIiI
            iIhkAIiIiGiIiP/4iIiIiIiIdgCIiIhoiIj/iIiIiIiIiIYAiIiIeIiIiIiIiIiIiIiGAAiIiIaP///
            ////////4hgAAAAAGZmZmZmZmZmZmZmYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD////////////////gA
            AAf4AAAD+AAAAfgAAAHAAAABwAAAAcAAAAHAAAAAwAAAAMAAAADAAAAAwAAAAMAAAABAAAAAQAAAAEA
            AAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAP4AAAH//////////////////////////w==

            Description
            -----------
            Returns the Base64 encoded representation of the icon

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBase64 | Clip

            Description
            -----------
            Returns the Base64 encoded representation of the icon and saves it to the clipboard.

        .EXAMPLE
            (Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBytes) -Join ''

            0010103232162900002322002200040000320006400010400000128200000000000000000000000
            0128001280001281280128000128012801281280012812812801921921920002550025500025525
            5025500025502550255255002552552550000000000000000000000000000000000000000000000
            0000000000000000000000000000000000006102102102102102102102102102102960000613611
            9119119119119120136136136118000119102134102102102102102102102102102134640011810
            2118102102102102102102102102102119960011810211910210413510212013613611810210496
            0011810211910211125513513613613613410210496001181021031021031432481201361191021
            0210396001191021031021021042552481181021021021031180011810210410210210214325513
            5102102102103118001191021031181021021031432481181021021021340011911910311810210
            2102232255248102102102134001191191181351191181021101432551021021021180013511911
            8135119119136136255248102102102119960136119118136136136136143255136135118102119
            9601361361341361361361362552551361361361361369601361361351201361361432552481361
            3613613613696013613613610413613625525513613613613613613610001361361361041361362
            5524813613613613613613611801361361361041361362551361361361361361361361340136136
            1361201361361361361361361361361361361340813613613414325525525525525525525525524
            8134000061021021021021021021021021021021020000000000000000000000000000000000000
            0000000000000000000000000000000000000000000025525525525525525525525525525525525
            5224003122400152240072240070007000700070003000300030003000300010001000100010000
            0000000000000000000012800025400125525525525525525525525525525525525525525525525
            5255255255255

            Description
            -----------
            Returns the bytes representation of the icon. -Join was used in this for the sake
            of displaying all of the data.

    #>
    [cmdletbinding(
        DefaultParameterSetName = '__DefaultParameterSetName'
    )]
    Param (
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullorEmpty()]
        [string]$Path,
        [parameter(ParameterSetName = 'Bytes')]
        [switch]$ToBytes,
        [parameter(ParameterSetName = 'Bitmap')]
        [switch]$ToBitmap,
        [parameter(ParameterSetName = 'Base64')]
        [switch]$ToBase64
    )
    Begin {
        If ($PSBoundParameters.ContainsKey('Debug')) {
            $DebugPreference = 'Continue'
        }
        Add-Type -AssemblyName System.Drawing
    }
    Process {
        $Path = Convert-Path -Path $Path
        Write-Debug $Path
        If (Test-Path -Path $Path) {
            $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path)| 
            Add-Member -MemberType NoteProperty -Name FullName -Value $Path -PassThru
            If ($PSBoundParameters.ContainsKey('ToBytes')) {
                Write-Verbose "Retrieving bytes"
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $MemoryStream.ToArray()   
                $MemoryStream.Flush()  
                $MemoryStream.Dispose()           
            } ElseIf ($PSBoundParameters.ContainsKey('ToBitmap')) {
                $Icon.ToBitMap()
            } ElseIf ($PSBoundParameters.ContainsKey('ToBase64')) {
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $Bytes = $MemoryStream.ToArray()   
                $MemoryStream.Flush() 
                $MemoryStream.Dispose()
                [convert]::ToBase64String($Bytes)
            }  Else {
                $Icon
            }
        } Else {
            Write-Warning "$Path does not exist!"
            Continue
        }
    }
}

# End Get-Icon

# Get Mapped Drive
Function Get-MappedDrive {
	param (
	    [string]$computername = "localhost"
	)
	    Get-WmiObject -Class Win32_MappedLogicalDisk -ComputerName $computername | 
	    Format-List DeviceId, VolumeName, SessionID, Size, FreeSpace, ProviderName
	}
# End Get Mapped Drive

# LayZ LazyWinAdmin GUI Tool
Function LayZ {
    C:\LazyWinAdmin\LazyWinAdmin\LazyWinAdmin.ps1
    }
# End LayZ LazyWinAdmin GUI Tool

# User Last Login
Function Get-UserLastLogonTime{

<#
.SYNOPSIS
Gets the last logon time of users on a Computer.

.DESCRIPTION
Pulls information from the wmi object Win32_UserProfile and outputs an array of objects with properties Name and LastUseTime.
If a date that is year 1 is outputted, then an error occured.

.PARAMETER ComputerName
[object] Specify which computer to target when finding logged on Users.
Default is the host computer

.PARAMETER User
[string] Specify a user to find on the computer.

.PARAMETER ListAllUsers
[switch] Specify the Function to list all users that logged into the computer.

.PARAMETER GetLastUsers
[switch] Specify the Function to get the last user to log onto the computer.

.PARAMETER ListCommonUsers
[switch] Specify to the Function to list common user.

.INPUTS
You may pipe objects into the ComputerName parameter.

.OUTPUTS
outputs an object array with a size dependant on the number of users that logged in with propeties Name and LastUseTime.


#>

    [cmdletBinding()]
    param(
        #computer Name
        [parameter(Mandatory = $False, ValueFromPipeline = $True)]
        [object]$ComputerName = $env:COMPUTERNAME,

        #parameter set, can only choose one from this group
        [parameter(Mandatory = $False, parameterSetName = 'user')]
        [string] $User,
        [parameter(ParameterSetName = 'all users')]
        [switch] $ListAllUsers,
        [parameter(ParameterSetName = 'Last user')]
        [switch] $GetLastUser,

        #Whether or not you want the Function to list Common users
        [switch] $ListCommonUsers
    )

    #Begin Pipeline
    Begin{
        #List of users that are present on all PCs, these won't output unless specified to do so.
        $CommonUsers = ("NetworkService", "LocalService", "systemprofile")
    }
    
    #Process Pipeline
    Process{
        #ping the machine before trying to do anything
        if(Test-Connection $ComputerName -Count 2 -Quiet){
            #try to get the OS version of the computer
            try{$OS = (gwmi  -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction stop).caption}
            catch{
                #had an error getting the WMI-object
                return New-Object psobject -Property @{
                            User = "Error getting WMIObject Win32_OperatingSystem"
                            LastUseTime = get-date 0
                            }
              }
            #make sure the OS retrieved is either Windows 7 or Windows 10 as this Function has not been set to work on other operating systems
            if($OS.contains("Windows 10") -or $OS.Contains("Windows 7")){
                try{
                    #try to get the WMiObject win32_UserProfile
                    $UserObjects = Get-WmiObject -ComputerName $ComputerName -Class Win32_userProfile -Property LocalPath,LastUseTime -ErrorAction Stop
                    $users = @()
                    #loop that handles all the data that came back from the WMI objects
                    forEach($UserObject in $UserObjects){
                        #extract the username from the local path, first find where the last slash is after, get everything after that into its own string
                        $i = $UserObject.localPath.LastIndexOf("\") + 1
                        $tempUserString = ""
                        while($UserObject.localPath.toCharArray()[$i] -ne $null){
                            $tempUserString += $UserObject.localPath.toCharArray()[$i]
                            $i++
                        }
                        #if list common users is turned on, skip this block
                        if(!$listCommonUsers){
                            #check if the user extracted from the local path is a common user
                            $isCommonUser = $false
                            forEach($userName in $CommonUsers){ 
                                if($userName -eq $tempUserString){
                                    $isCommonUser = $true
                                    break
                                }
                            }
                        }
                        #if the user is one of the users specified in the common users, skip it unless otherwise specified
                        if($isCommonUser){continue}
                        #check to see if the user has a timestamp for there last logon 
                        if($UserObject.LastUseTime -ne $null){
                            #This converts the string timestamp to a DateTime object
                            $TempUserLastUseTime = ([WMI] '').ConvertToDateTime($userObject.LastUseTime)
                        }
                        #the user had a local path but no timestamp was found for last logon
                        else{$TempUserLastUseTime = Get-Date 0}
                        #add user to array
                        $users += New-Object -TypeName psobject -Property @{
                            User = $tempUserString
                            LastUseTime = $TempUserLastUseTime
                            }
                    }
                }
                catch{
                    #error trying to retrieve WMI obeject win32_userProfile
                    return New-Object psobject -Property @{
                        User = "Error getting WMIObject Win32_userProfile"
                        LastUseTime = get-date 0
                        }
                }
            }
            else{
                #OS version was not compatible
                return New-Object psobject -Property @{
                    User = "Operating system $OS is not compatible with this Function."
                    LastUseTime = get-date 0
                    }
            }
        }
        else{
            #Computer was not pingable
            return New-Object psobject -Property @{
                User = "Can't Ping"
                LastUseTime = get-date 0
                }
        }

        #check to see if any users came out of the main Function
        if($users.count -eq 0){
            $users += New-Object -TypeName psobject -Property @{
                User = "NoLoggedOnUsers"
                LastUseTime = get-date 0
            }
        }
        #sort the user array by the last time they logged in
        else{$users = $users | Sort-Object -Property LastUseTime -Descending}
        #main output block
        #if List all users was chosen, output the full list of users found
        if($ListAllUsers){return $users}
        #if get last user was chosen, output the last user to log on the computer
        elseif($GetLastUser){return ($users[0])}
        else{
            #see if the user specified ever logged on
            ForEach($Username in $users){
                if($Username.User -eq $user) {return ($Username)}            
            }

            #user did not log on
            return New-Object psobject -Property @{
                User = "$user"
                LastUseTime = get-date 0
                }
        }
    }
    #End Pipeline
    End{Write-Verbose "Function get-UserLastLogonTime is complete"}
}
# End User Last Login

# Begin Unblock
Function Unblock ($path) { 

Get-ChildItem "$path" -Recurse | Unblock-File

}
# End Unblock 
Function Add-Help
{
 $helpText = @"
<#
.SYNOPSIS
    What does this do? 

.PARAMETER param1 
    What is param1?

.PARAMETER param2 
    What is param2?

.NOTES
    NAME: $($psISE.CurrentFile.DisplayName)
    AUTHOR: $env:username
    LASTEDIT: $(Get-Date)
    KEYWORDS:

.LINK
    http://julianscorner.com

.EXAMPLE     
    '12345' | THIS-Function -param1 180   
    Describe what this example accomplishes 
       
.EXAMPLE     
    THIS-Function -param2 @("text1","text2") -param1 180   
    Describe what this example accomplishes 

#Requires -Version 2.0
#>

"@
 $psise.CurrentFile.Editor.InsertText($helpText)
}

Function Add-FunctionTemplate
{
  $text1 = @"
Function THIS-Function
{
"@
  $text2 = @"
	[CmdletBinding()]
  param
  (
    [Parameter(Mandatory = `$true,
               ValueFromPipeline = `$true)]   
    [array]`$param1,                                   
    [Parameter(Mandatory = `$true)]   
    [int]`$param2
  )   
  BEGIN
  {
    # This block is used to provide optional one-time pre-processing for the Function.
    # PowerShell uses the code in this block one time for each instance of the Function in the pipeline.
  }
  PROCESS
  {
    # This block is used to provide record-by-record processing for the Function.
    # This block might be used any number of times, or not at all, depending on the input to the Function.
    # For example, if the Function is the first command in the pipeline, the Process block will be used one time.
    # If the Function is not the first command in the pipeline, the Process block is used one time for every
    # input that the Function receives from the pipeline.
    # If there is no pipeline input, the Process block is not used.
  }
  END
  {
    # This block is used to provide optional one-time post-processing for the Function.
  } 
}
"@
 $psise.CurrentFile.Editor.InsertText($text1)
 Add-Help
 $psise.CurrentFile.Editor.InsertText($text2)
}

Function Remove-AliasFromScript
{
  Get-Alias | 
    Select-Object Name, Definition | 
    ForEach-Object -Begin { $a = @{} } -Process {$a.Add($_.Name, $_.Definition)} -End {}

  $b = $errors = $null
  $b = $psISE.CurrentFile.Editor.Text

  [system.management.automation.psparser]::Tokenize($b,[ref]$errors) |
    Where-Object { $_.Type -eq "command" } |
      ForEach-Object `
      {
        if ($a.($_.Content))
        {
          $b = $b -replace
            ('(?<=(\W|\b|^))' + [regex]::Escape($_.content) + '(?=(\W|\b|$))'),
              $a.($_.content)
        }
      }

  $ScriptWithoutAliases = $psISE.CurrentPowerShellTab.Files.Add()
  $ScriptWithoutAliases.Editor.Text = $b
  $ScriptWithoutAliases.Editor.SetCaretPosition(1,1)
  $ScriptWithoutAliases.Editor.EnsureVisible(1)  
}

Function Replace-SpacesWithTabs
{
  param
  (
    [int]$spaces = 2
  ) 
  
  $tab = "`t"
  $space = " " * $spaces
  $text = $psISE.CurrentFile.Editor.Text

  $newText = ""
  
  foreach ($line in $text -split [Environment]::NewLine)
  {
    if ($line -match "\S")
    {
      $pos = $line.IndexOf($Matches[0])
      $indentation = $line.SubString(0, $pos)
      $remainder = $line.SubString($pos)
      
      $replaced = $indentation -replace $space, $tab
      
      $newText += $replaced + $remainder + [Environment]::NewLine
    }
    else
    {
      $newText += $line + [Environment]::NewLine
    }

    $psISE.CurrentFile.Editor.Text  = $newText
  }
}

Function Replace-TabsWithSpaces
{
  param
  (
    [int]$spaces = 2
  )   
  
  $tab = "`t"
  $space = " " * $spaces
  $text = $psISE.CurrentFile.Editor.Text

  $newText = ""
  
  foreach ($line in $text -split [Environment]::NewLine)
  {
    if ($line -match "\S")
    {
      $pos = $line.IndexOf($Matches[0])
      $indentation = $line.SubString(0, $pos)
      $remainder = $line.SubString($pos)
      
      $replaced = $indentation -replace $tab, $space
      
      $newText += $replaced + $remainder + [Environment]::NewLine
    }
    else
    {
      $newText += $line + [Environment]::NewLine
    }

    $psISE.CurrentFile.Editor.Text  = $newText
  }
}

Function Indent-SelectedText
{
  param
  (
    [int]$spaces = 2
  )
  
  $tab = " " * $space
  $text = $psISE.CurrentFile.Editor.SelectedText

  $newText = ""
  
  foreach ($line in $text -split [Environment]::NewLine)
  {
    $newText += $tab + $line + [Environment]::NewLine
  }

   $psISE.CurrentFile.Editor.InsertText($newText)
}

Function Add-RemarkedText
{
<#
.SYNOPSIS
    This Function will add a remark character # to selected text in the ISE.
    These are comment characters, and is great when you want to comment out
    a section of PowerShell code.

.NOTES
    NAME:  Add-RemarkedText
    AUTHOR: ed wilson, msft
    LASTEDIT: 05/16/2013
    KEYWORDS: Windows PowerShell ISE, Scripting Techniques

.LINK
     http://www.ScriptingGuys.com

#Requires -Version 2.0
#>
  $text = $psISE.CurrentFile.Editor.SelectedText

  foreach ($l in $text -Split [Environment]::NewLine)
  {
   $newText += "{0}{1}" -f ("#" + $l),[Environment]::NewLine
  }

  $psISE.CurrentFile.Editor.InsertText($newText)
}

Function Remove-RemarkedText
{
<#
.SYNOPSIS
    This Function will remove a remark character # to selected text in the ISE.
    These are comment characters, and is great when you want to clean up a
    previously commentted out section of PowerShell code.

.NOTES
    NAME:  Add-RemarkedText
    AUTHOR: ed wilson, msft
    LASTEDIT: 05/16/2013
    KEYWORDS: Windows PowerShell ISE, Scripting Techniques

.LINK
     http://www.ScriptingGuys.com

#Requires -Version 2.0
#>
  $text = $psISE.CurrentFile.Editor.SelectedText

  foreach ($l in $text -Split [Environment]::NewLine)
  {
    $newText += "{0}{1}" -f ($l -Replace '#',''),[Environment]::NewLine
  }

  $psISE.CurrentFile.Editor.InsertText($newText)
}
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

########################
# Office Word Functions#
########################
Function Add-OSCPicture
{
<#
.SYNOPSIS
Add-OSCPicture is an advanced Function which can be used to insert many pictures into a word document.
.DESCRIPTION
Add-OSCPicture is an advanced Function which can be used to insert many pictures into a word document.
.PARAMETER  <Path>
Specifies the path of slide.
.EXAMPLE
C:\PS> Add-OSCPicture -WordDocumentPath D:\Word\Document.docx -ImageFolderPath "C:\Users\Public\Pictures\Sample Pictures"
Action(Insert) ImageName
-------------- ---------
Finished   Chrysanthemum.jpg
Finished   Desert.jpg
Finished   Hydrangeas.jpg
Finished   Jellyfish.jpg
Finished   Koala.jpg
Finished   Lighthouse.jpg
Finished   Penguins.jpg
Finished   Tulips.jpg

This command shows how to insert many pictures to word document.
#>
[CmdletBinding()]
    Param(
    [Parameter(Mandatory=$true,Position=0)]
    [Alias('wordpath')]
    [String]$WordDocumentPath,
    [Parameter(Mandatory=$true,Position=1)]
    [Alias('imgpath')]
    [String]$ImageFolderPath
    )

If(Test-Path -Path $WordDocumentPath)
{
    If(Test-Path -Path $ImageFolderPath)
    {
    $WordExtension = (Get-Item -Path $WordDocumentPath).Extension
    If($WordExtension -like ".doc" -or $WordExtension -like ".docx")
        {
    $ImageFiles = Get-ChildItem -Path $ImageFolderPath -Recurse -Include *.emf,*.wmf,*.jpg,*.jpeg,*.jfif,*.png,*.jpe,*.bmp,*.dib,*.rle,*.gif,*.emz,*.wmz,*.pcz,*.tif,*.tiff,*.eps,*.pct,*.pict,*.wpg

    If($ImageFiles)
    {
    #Create the Word application object
    $WordAPP = New-Object -ComObject Word.Application
    $WordDoc = $WordAPP.Documents.Open("$WordDocumentPath")

    Foreach($ImageFile in $ImageFiles)
    {
    $ImageFilePath = $ImageFile.FullName

    $Properties = @{'ImageName' = $ImageFile.Name
    'Action(Insert)' = Try
    {
    $WordAPP.Selection.EndKey(6)|Out-Null
    $WordApp.Selection.InlineShapes.AddPicture("$ImageFilePath")|Out-Null
    $WordApp.Selection.InsertNewPage() #insert new page to word
    "Finished"
    }
    Catch
    {
    "Unfinished"
    }
    }

    $objWord = New-Object -TypeName PSObject -Property $Properties
    $objWord
    }

    $WordDoc.Save()
    $WordDoc.Close()
    $WordAPP.Quit()#release the object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordAPP)|Out-Null
    Remove-Variable WordAPP
    }
    Else
    {
    Write-Warning "There is no image in this '$ImageFolderPath' folder."
    }
    }
    Else
    {
    Write-Warning "There is no word document file in this '$WordDocumentPath' folder."
    }
    }
    Else
    {
    Write-Warning "Cannot find path '$ImageFolderPath' because it does not exist."
    }
    }
    Else
    {
    Write-Warning "Cannot find path '$WordDocumentPath' because it does not exist."
    }
    }


Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Columns = $null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Headers = $null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines=$false,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Format = '-231'
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Columns -eq $null) -and ($Headers -ne $null)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $null;
		}
		ElseIf(($Columns -ne $null) -and ($Headers -ne $null)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
        [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Columns -eq $null) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Columns -eq $null) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end Switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $true);
			$ConvertToTableArguments.Add("ApplyShading", $true);
			$ConvertToTableArguments.Add("ApplyFont", $true);
			$ConvertToTableArguments.Add("ApplyColor", $true);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $true); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $true);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $true);
			$ConvertToTableArguments.Add("ApplyLastColumn", $true);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$null,                                          # Modifiers
			$null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting)
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		#the next line causes the heading row to flow across page breaks
		$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

########################
# AD Functions##########
###############

Function Get-ADFSMORole
{
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

Function Get-AccountLockedOut
{
	
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

Function Get-ADGPOReplication
{
	<#
	.SYNOPSIS
		This Function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.DESCRIPTION
		This Function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.PARAMETER GPOName
		Specify the name of the GPO
	.PARAMETER All
		Specify that you want to retrieve all the GPO (slow if you have a lot of Domain Controllers)
	.EXAMPLE
		Get-ADGPOReplication -GPOName "Default Domain Policy"
	.EXAMPLE
		Get-ADGPOReplication -All
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		lazywinadmin.com
	
		VERSION HISTORY
		1.0 2014.09.22 	Initial version
						Adding some more Error Handling
						Fix some typo
	#>
	#requires -version 3
	[CmdletBinding()]
	PARAM (
		[parameter(Mandatory = $True, ParameterSetName = "One")]
		[String[]]$GPOName,
		[parameter(Mandatory = $True, ParameterSetName = "All")]
		[Switch]$All
	)
	BEGIN
	{
		TRY
		{
			if (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction Stop -ErrorVariable ErrorBeginIpmoAD }
			if (-not (Get-Module -Name GroupPolicy)) { Import-Module -Name GroupPolicy -ErrorAction Stop -ErrorVariable ErrorBeginIpmoGP }
		}
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			IF ($ErrorBeginIpmoAD) { Write-Warning -Message "[BEGIN] Error while Importing the module Active Directory" }
			IF ($ErrorBeginIpmoGP) { Write-Warning -Message "[BEGIN] Error while Importing the module Group Policy" }
			Write-Warning -Message "[BEGIN] $($Error[0].exception.message)"
		}
	}
	PROCESS
	{
		FOREACH ($DomainController in ((Get-ADDomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetDC -filter *).hostname))
		{
			TRY
			{
				IF ($psBoundParameters['GPOName'])
				{
					Foreach ($GPOItem in $GPOName)
					{
						$GPO = Get-GPO -Name $GPOItem -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPO
						
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPOItem
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}#Foreach ($GPOItem in $GPOName)
				}#IF ($psBoundParameters['GPOName'])
				IF ($psBoundParameters['All'])
				{
					$GPOList = Get-GPO -All -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPOAll
					
					foreach ($GPO in $GPOList)
					{
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPO.DisplayName
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}
				}#IF ($psBoundParameters['All'])
			}#TRY
			CATCH
			{
				Write-Warning -Message "[PROCESS] Something wrong happened"
				IF ($ErrorProcessGetDC) { Write-Warning -Message "[PROCESS] Error while running retrieving Domain Controllers with Get-ADDomainController" }
				IF ($ErrorProcessGetGPO) { Write-Warning -Message "[PROCESS] Error while running Get-GPO" }
				IF ($ErrorProcessGetGPOAll) { Write-Warning -Message "[PROCESS] Error while running Get-GPO -All" }
				Write-Warning -Message "[PROCESS] $($Error[0].exception.message)"
			}
		}#FOREACH
	}#PROCESS
}

Function Get-NestedMember
{
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

Function Get-ParentGroup
{
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
        [Parameter(Mandatory = $true)]
        [String[]]$Name
    )
    BEGIN 
    {
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
            FOREACH ($Obj in $Name)
            {
                # Make an Ambiguous Name Resolution
                $ADObject = Get-ADObject -LDAPFilter "(|(anr=$obj)(distinguishedname=$obj))" -Properties memberof -ErrorAction Stop
                IF($ADObject)
                {
                    # Show a warning if more than 1 object is found
                    if ($ADObject.count -gt 1){Write-Warning -Message "More than one object found with the $obj request"}
                    
                    FOREACH ($Account in $ADObject)
                    {
                        Write-Verbose -Message "[PROCESS] $($Account.name)"
                        $Account | Select-Object -ExpandProperty memberof | ForEach-Object -Process {

                            $CurrentObject = Get-Adobject -LDAPFilter "(|(anr=$_)(distinguishedname=$_))" -Properties Samaccountname
                                
                            
                            Write-Output $CurrentObject | Select-Object Name,SamAccountName,ObjectClass, @{L="Child";E={$Account.samaccountname}}
                            
                            Write-Verbose -Message "Inception - $($CurrentObject.distinguishedname)"
                            Get-ParentGroup -OutBuffer $CurrentObject.distinguishedname

                        }#$Account | Select-Object
                    }#FOREACH ($Account in $ADObject){
                }#IF($ADObject)
                ELSE {
                    #Write-Warning -Message "[PROCESS] Can't find the object $Obj"
                }#ELSE
            }#FOREACH ($Obj in $Object)
        }#TRY
        CATCH{
            Write-Warning -Message "[PROCESS] An Error occured"
            Write-Warning -Message $error[0].exception.message }
    }#PROCESS
    END
    {
        Write-Verbose -Message "[END] Get-NestedMember"
    }

###################
# End AD Functions#
###################
#######################
# Office 365 Functions#
#######################

Function Connect-Office365
{
<#
.SYNOPSIS
    This Function will prompt for credentials, load module MSOLservice,
	load implicit modules for Office 365 Services (AD, Lync, Exchange) using PSSession.
.DESCRIPTION
    This Function will prompt for credentials, load module MSOLservice,
	load implicit modules for Office 365 Services (AD, Lync, Exchange) using PSSession.
.EXAMPLE
    Connect-Office365
   
    This will prompt for your credentials and connect to the Office365 services
.EXAMPLE
    Connect-Office365 -verbose
   
    This will prompt for your credentials and connect to the Office365 services.
	Additionally you will see verbose messages on the screen to follow what is happening in the background
.NOTE
    Francois-Xavier Cat
    lazywinadmin.com
    @lazywinadm
#>
	[CmdletBinding()]
	PARAM ()
	BEGIN
	{
		TRY
		{
			#Modules
			IF (-not (Get-Module -Name MSOnline -ListAvailable))
			{
				Write-Verbose -Message "BEGIN - Import module Azure Active Directory"
				Import-Module -Name MSOnline -ErrorAction Stop -ErrorVariable ErrorBeginIpmoMSOnline
			}
			
			IF (-not (Get-Module -Name LyncOnlineConnector -ListAvailable))
			{
				Write-Verbose -Message "BEGIN - Import module Lync Online"
				Import-Module -Name LyncOnlineConnector -ErrorAction Stop -ErrorVariable ErrorBeginIpmoLyncOnline
			}
		}
		CATCH
		{
			Write-Warning -Message "BEGIN - Something went wrong!"
			IF ($ErrorBeginIpmoMSOnline)
			{
				Write-Warning -Message "BEGIN - Error while importing MSOnline module"
			}
			IF ($ErrorBeginIpmoLyncOnline)
			{
				Write-Warning -Message "BEGIN - Error while importing LyncOnlineConnector module"
			}
			
			Write-Warning -Message $error[0].exception.message
		}
	}
	PROCESS
	{
		TRY
		{
			
			# CREDENTIAL
			Write-Verbose -Message "PROCESS - Ask for Office365 Credential"
			$O365cred = Get-Credential -ErrorAction Stop -ErrorVariable ErrorCredential
			
			# AZURE ACTIVE DIRECTORY (MSOnline)
			Write-Verbose -Message "PROCESS - Connect to Azure Active Directory"
			Connect-MsolService -Credential $O365cred -ErrorAction Stop -ErrorVariable ErrorConnectMSOL
			
			# EXCHANGE ONLINE
			Write-Verbose -Message "PROCESS - Create session to Exchange online"
			$ExchangeURL = "https://ps.outlook.com/powershell/"
			$O365PS = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeURL -Credential $O365cred -Authentication Basic -AllowRedirection -ErrorAction Stop -ErrorVariable ErrorConnectExchange
			
			Write-Verbose -Message "PROCESS - Open session to Exchange online (Prefix: Cloud)"
			Import-PSSession -Session $O365PS –Prefix ExchCloud
			
			# LYNC ONLINE (LyncOnlineConnector)
			Write-Verbose -Message "PROCESS - Create session to Lync online"
			$lyncsession = New-CsOnlineSession –Credential $O365cred -ErrorAction Stop -ErrorVariable ErrorConnectExchange
			Import-PSSession -Session $lyncsession -Prefix LyncCloud
			
			# SHAREPOINT ONLINE
			#Connect-SPOService -Url https://contoso-admin.sharepoint.com –credential $O365cred
		}
		CATCH
		{
			Write-Warning -Message "PROCESS - Something went wrong!"
			IF ($ErrorCredential)
			{
				Write-Warning -Message "PROCESS - Error while gathering credential"
			}
			IF ($ErrorConnectMSOL)
			{
				Write-Warning -Message "PROCESS - Error while connecting to Azure AD"
			}
			IF ($ErrorConnectExchange)
			{
				Write-Warning -Message "PROCESS - Error while connecting to Exchange Online"
			}
			IF ($ErrorConnectLync)
			{
				Write-Warning -Message "PROCESS - Error while connecting to Lync Online"
			}
			
			Write-Warning -Message $error[0].exception.message
		}
	}
}
Function New-CimSmartSession
{
<# 
.SYNOPSIS 
    Function to create a CimSession to remote computer using either WSMAN or DCOM protocol.
	
.DESCRIPTION 
    Function to create a CimSession to remote computer using either WSMAN or DCOM protocol.
	This Function requires at least PowerShell v3.
	
.PARAMETER ComputerName 
    Specifies the ComputerName 
	
.PARAMETER Credential 
    Specifies alternative credentials
	
.EXAMPLE 
    New-CimSmartSession -ComputerName DC01,DC02
	
.EXAMPLE 
    $Session = New-CimSmartSession -ComputerName DC01 -Credential (Get-Credential -Credential "FX\SuperAdmin")
	New-CimInstance -CimSession $Session -Class Win32_Bios
	
.NOTES 
    Francois-Xavier Cat
	lazywinadmin.com
	@lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline = $true)]
		[string[]]$ComputerName = $env:COMPUTERNAME,
		
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)
	
	BEGIN
	{
		# Default Verbose/Debug message
		Function Get-DefaultMessage
		{
	<#
	.SYNOPSIS
		Helper Function to show default message used in VERBOSE/DEBUG/WARNING
	.DESCRIPTION
		Helper Function to show default message used in VERBOSE/DEBUG/WARNING.
		Typically called inside another Function in the BEGIN Block
	#>
			PARAM ($Message)
			Write-Output "[$(Get-Date -Format 'yyyy/MM/dd-HH:mm:ss:ff')][$((Get-Variable -Scope 1 -Name MyInvocation -ValueOnly).MyCommand.Name)] $Message"
		}#Get-DefaultMessage
		
		# Create a containter (hashtable) for the properties (Splatting)
		$CIMSessionSplatting = @{ }
			
		# Credential specified
		IF ($PSBoundParameters['Credential']) { $CIMSessionSplatting.Credential = $Credential }
		
		# CIMSession Option for DCOM (Default is WSMAN)
		$CIMSessionOption =	New-CimSessionOption -Protocol Dcom
	}
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Test-Connection")
			IF (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
			{
				$CIMSessionSplatting.ComputerName = $Computer
				
				
				# WSMAN Protocol
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: ([3-9]|[1-9][0-9]+)\.[0-9]+')
				{
					TRY
					{
						#WSMAN (Default when using New-CimSession)
						Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Connecting using WSMAN protocol (Default, requires at least PowerShell v3.0)")
						New-CimSession @CIMSessionSplatting -errorVariable ErrorProcessNewCimSessionWSMAN
					}
					CATCH
					{
						IF ($ErrorProcessNewCimSessionWSMAN) { Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Can't Connect using WSMAN protocol") }
						Write-Warning -Message (Get-DefaultMessage -Message $Error.Exception.Message)
					}
				}
				
				ELSE
				{
					# DCOM Protocol
					$CIMSessionSplatting.SessionOption = $CIMSessionOption
					
					TRY
					{
						Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Connecting using DCOM protocol")
						New-CimSession @SessionParams -errorVariable ErrorProcessNewCimSessionDCOM
					}
					CATCH
					{
						IF ($ErrorProcessNewCimSessionDCOM) { Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Can't connect using DCOM protocol either") }
						Write-Warning -Message (Get-DefaultMessage -Message $Error.Exception.Message)
					}
					FINALLY
					{
						# Remove the CimSessionOption for the DCOM protocol for the next computer
						$CIMSessionSplatting.Remove('CIMSessionOption')
					}
				}#ELSE
			}#Test-Connection
		}#FOREACH
	}#PROCESS
}#Function
Function New-DjoinFile
{
    <#
    .SYNOPSIS
        Function to generate a blob file accepted by djoin.exe tool (offline domain join)
    
	.DESCRIPTION
        Function to generate a blob file accepted by djoin.exe tool (offline domain join)
	
		This Function can create a file compatible with djoin with the Blob initially provisionned.
	
	.PARAMETER Blob
		Specifies the blob generated by djoin
	
	.PARAMETER DestinationFile
		Specifies the full path of the file that will be created
	
		Default is c:\temp\djoin.tmp
    
	.EXAMPLE
        New-DjoinFile -Blob $Blob -DestinationFile C:\temp\test.tmp
    
	.NOTES
        Francois-Xavier.Cat
        LazyWinAdmin.com
        @lazywinadm
        github.com/lazywinadmin
	
    .LINK
		https://github.com/lazywinadmin/PowerShell/tree/master/TOOL-New-DjoinFile
	.LINK
		http://www.lazywinadmin.com/2016/07/offline-domain-join-copying-djoin.html
	.LINK
        https://msdn.microsoft.com/en-us/library/system.io.fileinfo(v=vs.110).aspx
    #>
	[Cmdletbinding()]
	PARAM (
		[Parameter(Mandatory = $true)]
		[System.String]$Blob,
		[Parameter(Mandatory = $true)]
		[System.IO.FileInfo]$DestinationFile = "c:\temp\djoin.tmp"
	)
	
	PROCESS
	{
		TRY
		{
			# Create a byte object
			$bytechain = New-Object -TypeName byte[] -ArgumentList 2
			# Add the first two character for Unicode Encoding
			$bytechain[0] = 255
			$bytechain[1] = 254
			
			# Creates a write-only FileStream
			$FileStream = $DestinationFile.Openwrite()
			
			# Append Hash as byte
			$bytechain += [System.Text.Encoding]::unicode.GetBytes($Blob)
			# Append two extra 0 bytes characters
			$bytechain += 0
			$bytechain += 0
			
			# Write back to the file
			$FileStream.write($bytechain, 0, $bytechain.Length)
			
			# Close the file Stream
			$FileStream.Close()
		}
		CATCH
		{
			$Error[0]
		}
	}
}
Function New-IsoFile  
{  
  <#  
   .Synopsis  
    Creates a new .iso file  
   .Description  
    The New-IsoFile cmdlet creates a new .iso file containing content from chosen folders  
   .Example  
    New-IsoFile "c:\tools","c:Downloads\utils"  
    This command creates a .iso file in $env:temp folder (default location) that contains c:\tools and c:\downloads\utils folders. The folders themselves are included at the root of the .iso image.  
   .Example 
    New-IsoFile -FromClipboard -Verbose 
    Before running this command, select and copy (Ctrl-C) files/folders in Explorer first.  
   .Example  
    dir c:\WinPE | New-IsoFile -Path c:\temp\WinPE.iso -BootFile "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\efisys.bin" -Media DVDPLUSR -Title "WinPE" 
    This command creates a bootable .iso file containing the content from c:\WinPE folder, but the folder itself isn't included. Boot file etfsboot.com can be found in Windows ADK. Refer to IMAPI_MEDIA_PHYSICAL_TYPE enumeration for possible media types: http://msdn.microsoft.com/en-us/library/windows/desktop/aa366217(v=vs.85).aspx  
   .Notes 
    NAME:  New-IsoFile  
    AUTHOR: Chris Wu 
    LASTEDIT: 03/23/2016 14:46:50  
 #>  
  
  [CmdletBinding(DefaultParameterSetName='Source')]Param( 
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true, ParameterSetName='Source')]$Source,  
    [parameter(Position=2)][string]$Path = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",  
    [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})][string]$BootFile = $null, 
    [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER', 
    [string]$Title = (Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
    [switch]$Force, 
    [parameter(ParameterSetName='Clipboard')][switch]$FromClipboard 
  ) 
 
  Begin {  
    ($cp = new-object System.CodeDom.Compiler.CompilerParameters).CompilerOptions = '/unsafe' 
    if (!('ISOFile' -as [type])) {  
      Add-Type -CompilerParameters $cp -TypeDefinition @' 
public class ISOFile  
{ 
  public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks)  
  {  
    int bytes = 0;  
    byte[] buf = new byte[BlockSize];  
    var ptr = (System.IntPtr)(&bytes);  
    var o = System.IO.File.OpenWrite(Path);  
    var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;  
  
    if (o != null) { 
      while (TotalBlocks-- > 0) {  
        i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);  
      }  
      o.Flush(); o.Close();  
    } 
  } 
}  
'@  
    } 
  
    if ($BootFile) { 
      if('BDR','BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
      ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type=1}).Open()  # adFileTypeBinary 
      $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
      ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
    } 
 
    $MediaType = @('UNKNOWN','CDROM','CDR','CDRW','DVDROM','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','HDDVDROM','HDDVDR','HDDVDRAM','BDROM','BDR','BDRE') 
 
    Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))" 
    ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName=$Title}).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 
  
    if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break } 
  }  
 
  Process { 
    if($FromClipboard) { 
      if($PSVersionTable.PSVersion.Major -lt 5) { Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break } 
      $Source = Get-Clipboard -Format FileDropList 
    } 
 
    foreach($item in $Source) { 
      if($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) { 
        $item = Get-Item -LiteralPath $item 
      } 
 
      if($item) { 
        Write-Verbose -Message "Adding item to the target image: $($item.FullName)" 
        try { $Image.Root.AddTree($item.FullName, $true) } catch { Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') } 
      } 
    } 
  } 
 
  End {  
    if ($Boot) { $Image.BootImageOptions=$Boot }  
    $Result = $Image.CreateResultImage()  
    [ISOFile]::Create($Target.FullName,$Result.ImageStream,$Result.BlockSize,$Result.TotalBlocks) 
    Write-Verbose -Message "Target image ($($Target.FullName)) has been created" 
    $Target 
  } 
} 
Function New-Password
{
<#
	.SYNOPSIS
		Function to Generate a new password.
	
	.DESCRIPTION
		Function to Generate a new password.
		By default it will generate a 12 characters length password, you can change this using the parameter Length.
		I excluded the following characters: ",',.,/,1,<,>,`,O,0,l,|
        Some of those are ambiguous characters like 1 or l or |
		You can add exclusion by checking the following ASCII Table http://www.asciitable.com/
	
		If the length requested is less or equal to 4, it will generate a random password.
		If the length requested is greater than 4, it will make sure the password contains an Upper and Lower case letter, a Number and a special character
	
	.PARAMETER Length
		Specifies the length of the password.
        Default is 12 characters

    .PARAMETER Count
        Specifies how many password you want to output.
        Default is 1 password.
	
	.EXAMPLE
		PS C:\> New-Password -Length 30
		
		=E)(72&:f\W6:VRGE(,t1x6sZi-346

    .EXAMPLE
        PS C:\> New-Password 3
        
        !}R
	
	.NOTES
		See ASCII Table http://www.asciitable.com/
		Code based on a blog post of https://mjolinor.wordpress.com/2014/01/31/random-password-generator/
#>
	[CmdletBinding()]
	PARAM
	(
		[ValidateNotNull()]
		[int]$Length = 12,
        [ValidateRange(1,256)]
        [Int]$Count = 1
	)#PARAM
	
	BEGIN
	{
		# Create ScriptBlock with the ASCII Char Codes
		$PasswordCharCodes = { 33..126 }.invoke()
		
		
        # Exclude some ASCII Char Codes from the ScriptBlock
        #  Excluded characters are ",',.,/,1,<,>,`,O,0,l,|
		#  See http://www.asciitable.com/ for mapping
		34, 39, 46, 47, 49, 60, 62, 96, 48, 79, 108, 124 | ForEach-Object { [void]$PasswordCharCodes.Remove($_) }
		$PasswordChars = [char[]]$PasswordCharCodes
	}#BEGIN

	PROCESS
	{
        1..$count | ForEach-Object {
            # Password of 4 characters or longer
		    IF ($Length -gt 4)
		    {
			
			    DO
			    {
				    # Generate a Password of the length requested
				    $NewPassWord = $(foreach ($i in 1..$length) { Get-Random -InputObject $PassWordChars }) -join ''
			    }#Do
			    UNTIL (
			    # Make sure it contains an Upercase and Lowercase letter, a number and another special character
			    ($NewPassword -cmatch '[A-Z]') -and
			    ($NewPassWord -cmatch '[a-z]') -and
			    ($NewPassWord -imatch '[0-9]') -and
			    ($NewPassWord -imatch '[^A-Z0-9]')
			    )#Until
		    }#IF
            # Password Smaller than 4 characters
		    ELSE
		    {
			    $NewPassWord = $(foreach ($i in 1..$length) { Get-Random -InputObject $PassWordChars }) -join ''
		    }#ELSE
		
		    # Output a new password
		    Write-Output $NewPassword
        }
	} #PROCESS
	END
	{
        # Cleanup
		Remove-Variable -Name NewPassWord -ErrorAction 'SilentlyContinue'
	} #END
} #Function
Function New-RandomPassword
{
<#
.SYNOPSIS
	Function to generate a complex and random password
.DESCRIPTION
    Function to generate a complex and random password
	
	This is using the GeneratePassword method from the
	system.web.security.membership NET Class.
	
	https://msdn.microsoft.com/en-us/library/system.web.security.membership.generatepassword(v=vs.100).aspx
	
.PARAMETER Length
    The number of characters in the generated password. The length must be between 1 and 128 characters.
    Default is 12.

.PARAMETER NumberOfNonAlphanumericCharacters
    The minimum number of non-alphanumeric characters (such as @, #, !, %, &, and so on) in the generated password.
    Default is 5.

.PARAMETER Count
    Specifies how many password you want. Default is 1

.EXAMPLE
    New-RandomPassword
        []sHX@]W#w-{
.EXAMPLE
    New-RandomPassword -Length 8 -NumberOfNonAlphanumericCharacters 2
        v@Warq_6
.EXAMPLE
    New-RandomPassword -Count 5
        *&$6&d1[f8zF
        Ns$@[lRH{;f4
        ;G$Su^M$bS+W
        mgZ/{y8}I@-t
        **W.)60kY4$V
.NOTES
    francois-xavier.cat
    www.lazywinadmin.com
    @lazywinadm
    github.com/lazywinadmin
#>
	PARAM (
		[Int32]$Length = 12,
		
		[Int32]$NumberOfNonAlphanumericCharacters = 5,
		
		[Int32]$Count = 1
	)
	
	BEGIN
	{
		Add-Type -AssemblyName System.web;
	}
	
	PROCESS
	{
		1..$Count | ForEach-Object {
			[System.Web.Security.Membership]::GeneratePassword($Length, $NumberOfNonAlphanumericCharacters)
		}
	}
}
Function New-ScriptMessage
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
	
	.PARAMETER Block
		Specifies the Block where the message is coming from.
	
	.PARAMETER DateFormat
		Specifies the format of the date.
		Default is 'yyyy\/MM\/dd HH:mm:ss:ff' For example: 2016/04/20 23:33:46:78
	
	.PARAMETER FunctionScope
		Valid values are "Global", "Local", or "Script", or a number relative to the current scope (0 through the number of scopes, where 0 is the current scope and 1 is its parent). "Local" is the default
		
		See also: About_scopes https://technet.microsoft.com/en-us/library/hh847849.aspx
		
		Example:
		0 is New-ScriptMessage
		1 is the Function calling New-ScriptMessage
		2 is for example the script/Function calling the Function which call New-ScriptMessage
		etc...
	
	.EXAMPLE
		New-ScriptMessage -Message "Francois-Xavier" -Block PROCESS -Verbose -FunctionScope 0
		
		[2016/04/20 23:33:46:78][New-ScriptMessage][PROCESS] Francois-Xavier
	
	.EXAMPLE
		New-ScriptMessage -message "Connected"
		
		if the Function is just called from the prompt you will get the following output
		[2015/03/14 17:32:53:62] Connected
	
	.EXAMPLE
		New-ScriptMessage -message "Connected to $Computer" -FunctionScope 1
		
		If the Function is called from inside another Function,
		It will show the name of the Function.
		[2015/03/14 17:32:53:62][Get-Something] Connected
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
		github.com/lazywinadmin
#>
	
	[CmdletBinding()]
	[OutputType([string])]
	param
	(
		[String]$Message,
		[String]$Block,
		[String]$DateFormat = 'yyyy\/MM\/dd HH:mm:ss:ff',
		$FunctionScope = "1"
	)
	
	PROCESS
	{
		$DateFormat = Get-Date -Format $DateFormat
		$MyCommand = (Get-Variable -Scope $FunctionScope -Name MyInvocation -ValueOnly).MyCommand.Name
		IF ($MyCommand)
		{
			$String = "[$DateFormat][$MyCommand]"
		} #IF
		ELSE
		{
			$String = "[$DateFormat]"
		} #Else
		
		IF ($PSBoundParameters['Block'])
		{
			$String += "[$Block]"
		}
		Write-Output "$String $Message"
	} #Process
}
Function New-Shortcut
{
<#
.Synopsis
   Create file shortcut.
.DESCRIPTION
   You can create file shortcut into desired directory.
   Both Pipeline input and parameter input is supported.
.EXAMPLE
   New-Shortcut -TargetPaths "C:\Users\Administrator\Documents\hogehoge.csv" -Verbose -PassThru
    # Set Target full path in -TargetPaths (you can set multiple path). 
    # Set Directory to create shortcut in -ShortcutDirectory (Default is user Desktop).
    # Set -Verbose to sett Verbose status
    # Set -PassThru to output Shortcut creation result.

.NOTES
   Make sure file path is valid.
.COMPONENT
   COM
#>
    [CmdletBinding()]
    [OutputType([System.__ComObject])]
    param
    (
        # Set Target full path to create shortcut
        [parameter(
            position  = 0,
            mandatory = 1,
            ValueFromPipeline = 1,
            ValueFromPipeLineByPropertyName = 1)]
        [validateScript({$_ | %{Test-Path $_}})]
        [string[]]
        $TargetPaths,

        # set shortcut Directory to create shortcut. Default is user Desktop.
        [parameter(
            position  = 1,
            mandatory = 0,
            ValueFromPipeLineByPropertyName = 1)]
        [validateScript({-not(Test-Path $_)})]
        [string]
        $ShortcutDirectory = "$env:USERPROFILE\Desktop",

        # Set Description for shortcut.
        [parameter(
            position  = 2,
            mandatory = 0,
            ValueFromPipeLineByPropertyName = 1)]
        [string]
        $Description,

        # set if you want to show create shortcut result
        [parameter(
            position  = 3,
            mandatory = 0)]
        [switch]
        $PassThru
    )

    begin
    {
        $extension = ".lnk"
        $wsh = New-Object -ComObject Wscript.Shell
    }

    process
    {
        foreach ($TargetPath in $TargetPaths)
        {
            Write-Verbose ("Get filename from original target path '{0}'" -f $TargetPath)
            # Create File Name from original TargetPath
            $fileName = Split-Path $TargetPath -Leaf
            
            # set Path for Shortcut
            $path = Join-Path $ShortcutDirectory ($fileName + $extension)

            # Call Wscript to create Shortcut
            Write-Verbose ("Trying to create Shortcut for name '{0}'" -f $path)
            $shortCut = $wsh.CreateShortCut($path)
            $shortCut.TargetPath = $TargetPath
            $shortCut.Description = $Description
            $shortCut.Save()

            if ($PSBoundParameters.PassThru)
            {
                Write-Verbose ("Show Result for shortcut result for target file name '{0}'" -f $TargetPath)
                $shortCut
            }
        }
    }

    end
    {
    }
}
Function Out-DataTable {
    [CmdletBinding()]
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject)

    Begin
    {
    Function Get-Type {
        param($type)

        $types = @(
        'System.Boolean',
        'System.Byte[]',
        'System.Byte',
        'System.Char',
        'System.Datetime',
        'System.Decimal',
        'System.Double',
        'System.Guid',
        'System.Int16',
        'System.Int32',
        'System.Int64',
        'System.Single',
        'System.UInt16',
        'System.UInt32',
        'System.UInt64')

        if ( $types -contains $type ) {
            Write-Output "$type"
        }
        else {
            Write-Output 'System.String'
        
        }
    } #Get-Type
        $dt = new-object Data.datatable  
        $First = $true 
    }
    Process
    {
        foreach ($object in $InputObject)
        {
            $DR = $DT.NewRow()  
            foreach($property in $object.PsObject.get_properties())
            {  
                if ($first)
                {  
                    $Col =  new-object Data.DataColumn  
                    $Col.ColumnName = $property.Name.ToString()  
                    if ($property.value)
                    {
                        if ($property.value -isnot [System.DBNull]) {
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)")
                         }
                    }
                    $DT.Columns.Add($Col)
                }  
                if ($property.Gettype().IsArray) {
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1
                }  
               else {
                    If ($Property.Value) {
                        $DR.Item($Property.Name) = $Property.Value
                    } Else {
                        $DR.Item($Property.Name)=[DBNull]::Value
                    }
                }
            }  
            $DT.Rows.Add($DR)  
            $First = $false
        }
    } 
     
    End
    {
        Write-Output @(,($dt))
    }

}
Function Out-Excel
{
<#
	.SYNOPSIS
	.DESCRIPTION
	.PARAMETER Property
	.PARAMETER Raw
	.NOTES
    	Original Script: http://pathologicalscripter.wordpress.com/out-excel/
	
		TODO:
			Parameter to change color of header
			Parameter to activate background color on Odd unit
			Add TRY/CATCH
			Validate Excel first is present
#>
	[CmdletBinding()]
	PARAM ([string[]]$property, [switch]$raw)
	
	BEGIN
	{
		# start Excel and open a new workbook
		$Excel = New-Object -Com Excel.Application
		$Excel.visible = $True
		$Excel = $Excel.Workbooks.Add()
		$Sheet = $Excel.Worksheets.Item(1)
		# initialize our row counter and create an empty hashtable
		# which will hold our column headers
		$Row = 1
		$HeaderHash = @{ }
	}
	
	PROCESS
	{
		if ($_ -eq $null) { return }
		if ($Row -eq 1)
		{
			# when we see the first object, we need to build our header table
			if (-not $property)
			{
				# if we haven’t been provided a list of properties,
				# we’ll build one from the object’s properties
				$property = @()
				if ($raw)
				{
					$_.properties.PropertyNames | %{ $property += @($_) }
				}
				else
				{
					$_.PsObject.get_properties() | % { $property += @($_.Name.ToString()) }
				}
			}
			$Column = 1
			foreach ($header in $property)
			{
				# iterate through the property list and load the headers into the first row
				# also build a hash table so we can retrieve the correct column number
				# when we process each object
				$HeaderHash[$header] = $Column
				$Sheet.Cells.Item($Row, $Column) = $header.toupper()
				$Column++
			}
			# set some formatting values for the first row
			$WorkBook = $Sheet.UsedRange
			$WorkBook.Interior.ColorIndex = 19
			$WorkBook.Font.ColorIndex = 11
			$WorkBook.Font.Bold = $True
			$WorkBook.HorizontalAlignment = -4108
		}
		$Row++
		foreach ($header in $property)
		{
			# now for each object we can just enumerate the headers, find the matching property
			# and load the data into the correct cell in the current row.
			# this way we don’t have to worry about missing properties
			# or the “ordering” of the properties
			if ($thisColumn = $HeaderHash[$header])
			{
				if ($raw)
				{
					$Sheet.Cells.Item($Row, $thisColumn) = [string]$_.properties.$header
				}
				else
				{
					$Sheet.Cells.Item($Row, $thisColumn) = [string]$_.$header
				}
			}
		}
	}
	
	end
	{
		# now just resize the columns and we’re finished
		if ($Row -gt 1) { [void]$WorkBook.EntireColumn.AutoFit() }
	}
}
<#
.SYNOPSIS
  Outputs to a UTF-8-encoded file *without a BOM* (byte-order mark).

.DESCRIPTION
  Mimics the most important aspects of Out-File:
  * Input objects are sent to Out-String first.
  * -Append allows you to append to an existing file, -NoClobber prevents
    overwriting of an existing file.
  * -Width allows you to specify the line width for the text representations
     of input objects that aren't strings.
  However, it is not a complete implementation of all Out-String parameters:
  * Only a literal output path is supported, and only as a parameter.
  * -Force is not supported.

  Caveat: *All* pipeline input is buffered before writing output starts,
          but the string representations are generated and written to the target
          file one by one.

.NOTES
  The raison d'être for this advanced Function is that, as of PowerShell v5,
  Out-File still lacks the ability to write UTF-8 files without a BOM:
  using -Encoding UTF8 invariably prepends a BOM.

  Copyright (c) 2017 Michael Klement <mklement0@gmail.com> (http://same2u.net), 
  released under the [MIT license](https://spdx.org/licenses/MIT#licenseText).

#>
Function Out-FileUtf8NoBom {

  [CmdletBinding()]
  param(
    [Parameter(Mandatory, Position=0)] [string] $LiteralPath,
    [switch] $Append,
    [switch] $NoClobber,
    [AllowNull()] [int] $Width,
    [Parameter(ValueFromPipeline)] $InputObject
  )

  #requires -version 3

  # Make sure that the .NET framework sees the same working dir. as PS
  # and resolve the input path to a full path.
  [System.IO.Directory]::SetCurrentDirectory($PWD) # Caveat: .NET Core doesn't support [Environment]::CurrentDirectory
  $LiteralPath = [IO.Path]::GetFullPath($LiteralPath)

  # If -NoClobber was specified, throw an exception if the target file already
  # exists.
  if ($NoClobber -and (Test-Path $LiteralPath)) {
    Throw [IO.IOException] "The file '$LiteralPath' already exists."
  }

  # Create a StreamWriter object.
  # Note that we take advantage of the fact that the StreamWriter class by default:
  # - uses UTF-8 encoding
  # - without a BOM.
  $sw = New-Object IO.StreamWriter $LiteralPath, $Append

  $htOutStringArgs = @{}
  if ($Width) {
    $htOutStringArgs += @{ Width = $Width }
  }

  # Note: By not using begin / process / end blocks, we're effectively running
  #       in the end block, which means that all pipeline input has already
  #       been collected in automatic variable $Input.
  #       We must use this approach, because using | Out-String individually
  #       in each iteration of a process block would format each input object
  #       with an indvidual header.
  try {
    $Input | Out-String -Stream @htOutStringArgs | % { $sw.WriteLine($_) }
  } finally {
    $sw.Dispose()
  }

}

# If this script is invoked directly - as opposed to being dot-sourced in order
# to define the embedded Function for later use - invoke the embedded Function,
# relaying any arguments passed.
if (-not ($MyInvocation.InvocationName -eq '.' -or $MyInvocation.Line -eq '')) {
  Out-FileUtf8NoBom @Args
}

#Notes
<#
This script checks through a list of computers to report via email
whether any computers are in a "Reboot Pending" state.
#>

#Email Variables
$smtpServer = "smtp.contoso.com"
$smtpFrom = "Reboot Pending Report <Reboot.Pending@contoso.com>"
$smtpTo = "Server Admin <server.admin@contoso.com>"
$Subject = "'Reboot Pending' report"

#Server List
<#Three typical ways to get a list of computers.
1. Comma separated manually entered list.
$CommaList = ("Server01","Server02")
$List = $CommaList.Split(",")
2. List OU members 
$List = Get-ADComputer -Filter * -SearchBase "OU=IT,DC=contoso,DC=com"
3. List from TXT file
$List = Get-Content -Path "c:\scripts\PC-List.txt"
#>

# Choose a List method as shown above and replace the following two lines.
#$CommaList = ("Server01","Server02")
$CommaList = ("usonvsvritsw.USON.LOCAL")
$List = $CommaList.Split(",") 


<#~~~~~~~~~~~~~~~ DO NOT EDIT BEYOND THIS POINT ~~~~~~~~~~~~~~~#>

#Table Style
$style = "<style>BODY{font-family: Calibri; font-size: 10pt;}"
$style = $style + "TABLE{border: 0px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 0px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 0px solid black; padding: 5px; }"
$style = $style + "</style>"
$THstyle = "font-size:16pt;font-weight:bold;"
$TDstyle = "font-weight:bold;text-align:center;"

#Clean Start 
$Checks = ("WUboot,WUVal,PakBoot,PakVal,PakWowBoot,PakWowVal,RenFileBoot,PCnameBoot,CBSboot,CBSVal,Content")
$Clear = $checks.split(",")
Clear-Variable -Name $clear

#Table header row
$Content =""
$Content += "<table id=""t1"">"
$Content += "<tr bgcolor=#ADD8E6>"
$Content += "<td width=100 style=$THstyle> Server</td>"
$Content += "<td width=75 style=$TDstyle>Reboot Required?</td>"
$Content += "<td width=75 style=$TDstyle>Windows Updates</td>"
$Content += "<td width=75 style=$TDstyle>Package Installer</td>"
$Content += "<td width=75 style=$TDstyle>Package Installer 64</td>"
$Content += "<td width=75 style=$TDstyle>File Rename</td>"
$Content += "<td width=75 style=$TDstyle>Hostname Change</td>"
$Content += "<td width=80 style=$TDstyle>Component Based Svces</td>"


foreach ($PC in $CommaList)
#List from OU  - Change searchbase from Split to OU
#List from TXT
{
$Content += "<tr>"

#Windows Updates
$WUVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\" -Name RebootRequired -ErrorAction SilentlyContinue}) | foreach { $_.RebootRequired }
if ($WUVal -ne $null) {$WUBoot = "Yes"}
else {$WUBoot = "No"}

#Package Installer
$PakVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\Software\Microsoft\Updates\" -Name UpdateExeVolatile -ErrorAction SilentlyContinue}) | foreach { $_.UpdateExeVolatile }
if ($PakVal -ne $null) {$PakBoot = "Yes"}
else {$PakBoot = "No"}


#Package Installer - Wow64
$PakWowVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\Software\Wow6432Node\Microsoft\Updates\" -Name UpdateExeVolatile -ErrorAction SilentlyContinue}) | foreach { $_.UpdateExeVolatile }
if ($PakWowVal -ne $null) {$PakWowBoot = "Yes"}
else {$PakWowBoot = "No"}


#Pending File Rename Operation
$RenFileVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue}) | foreach { $_.PendingFileRenameOperations }
if ($RenFileVal -ne $null) {$RenFileBoot = "Yes"}
else {$RenFileBoot = "No"}


#Pending Computer Rename
$PCnameIs = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" -Name ComputerName -ErrorAction SilentlyContinue}) | foreach { $_.ComputerName }
$PCnameBe = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -Name ComputerName -ErrorAction SilentlyContinue}) | foreach { $_.ComputerName }
if ($PCnameIs -eq $PCnameBe) {$PCnameBoot = "No"}
else {$PCnameBoot = "Yes"}


#Component Based Servicing
$CBSVal = (Invoke-command -computer $PC {Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\" -Name RebootPending -ErrorAction SilentlyContinue}) | foreach { $_.RebootPending }
if ($CBSVal -ne $null) {$CBSBoot = "Yes"}
else {$CBSBoot = "No"}


#Email HTML Content - append loop
$Content += "<td bgcolor=#dddddd align=left><b>$PC</b></td>"
if (($WUboot,$PakBoot,$PakWowBoot,$RenFileBoot,$PCnameBoot,$CBSBoot) -contains "Yes")
{$Content += "<td bgcolor=#ff4000 align=center>Yes</td>"}
else
{$Content += "<td bgcolor=#65ff00 align=center>No</td>"}
$Content += "<td bgcolor=#f5f5f5 align=center>$WUboot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$PakBoot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$PakWowBoot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$RenFileBoot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$PCnameBoot</td>"
$Content += "<td bgcolor=#f5f5f5 align=center>$CBSBoot</td>"
$Content += "</tr>"
}

#Close HTML
$Content += "</table>"
$Content += "</body>"
$Content += "</html>"

#Send Email Report
#Send-MailMessage -From $smtpFrom -To $smtpTo -Subject $Subject -Body $Content -BodyAsHtml -Priority High -dno onSuccess, onFailure -SmtpServer $smtpServer
$Content | Out-File C:\LazyWinAdmin\PendingReboot.html

Function Query-SoftwareInstalled
{
[CmdletBinding (SupportsShouldProcess = $True)]
Param
(
    [Parameter (Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
        HelpMessage="Input a domain OU structure to query for software installed.`r`nE.g.- `"OU=1stOUName,OU=2ndOUName,DC=LowLevelDomain,DC=MidLevelDomain,DC=TopLevelDomain`"")]
    [Alias('OU')]
    [string]$OUStructure,

    [Parameter (Mandatory=$True,
    ValueFromPipeline=$False,
    ValueFromPipelineByPropertyName=$False,
        HelpMessage="Input the software name that is to be queried.`r`n(Be sure it matches the software's name listed in the registry.)")]
    [Alias('Install')]
    [string[]]$Software,

    [Parameter (Mandatory=$False,
    ValueFromPipeline=$False,
    ValueFromPipelineByPropertyName=$False,
        HelpMessage="Input the number of days before a computer account is considered 'Inactive'.`r`nE.g.- `"30`"")]
    [Alias('Days')]
    [int32]$InactivityThreshold = "30",

    [Parameter (Mandatory=$False,
    ValueFromPipeline=$False,
    ValueFromPipelineByPropertyName=$False,
        HelpMessage="Input a directory to output the CSV file.")]
    [Alias('Folder')]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\$Software-Machines"
)
$date = Get-Date -Format MMM-dd-yyyy
$time = [DateTime]::Now
$Computers = Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -SearchBase "$OUStructure" -Properties 'Name','OperatingSystem','CanonicalName','LastLogonTimeStamp'
If ($OutputPath)
{
    If (-not (Test-Path -Path $OutputPath))
    {
    New-Item -Path "$OutputPath" -ItemType Directory
    }
}
ForEach ($Computer in $Computers)
{
$subkeyarray = $null
$NetAdapterError = $null
$nocomp = $null
$comp = $null
$Name = $null
$IPAddress = $null
$OS = $null
$CanonicalName = $null
$DisplayName = $null
$DisplayVersion = $null
$InstallLocation = $null
$Publisher = $null
$NetConfig = $null
$MAC = $null
$IPEnabled = $null
$DNSServers = $null
$LogonTime = [DateTime]::FromFileTime($Computer.LastLogonTimeStamp)
    If ($LogonTime -gt (Get-Date).AddDays(-($InactivityThreshold)))
    {
    $Name = $($Computer.Name)
    $NetConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $($Computer.Name) -ErrorAction SilentlyContinue -ErrorVariable NetAdapterError | Where {$_.IPEnabled -eq $true}
        If ($NetAdapterError -like "*The RPC server is unavailable*")
        {
        $IPAddress = "ERROR: Remote connection to $($Computer.Name)`'s network adapter failed."
        }
        Else
        {
            ForEach ($AdapterItem in $NetConfig)
            {
            $MAC = $AdapterItem.MACAddress
            $IPAddress = $AdapterItem.IPAddress | Where {$_ -like "172.*"}
            $IPEnabled = $AdapterItem.IPEnabled
            $DNSServers = $AdapterItem.DNSServerSearchOrder
            }
        }
    $OS = $($Computer.OperatingSystem)
    $CanonicalName = $($Computer.CanonicalName)
    #Define the variable to hold the location of Currently Installed Programs
    $UninstallKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
    #Create an instance of the Registry Object and open the HKLM base key
    $subkeyarray = @()
    $PrevErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"
    Write-Error -Message "Test - Disregard"
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$($Computer.Name))
    $ErrorActionPreference = $PrevErrorActionPreference
        If ($error[0] -like "*Exception calling `"OpenRemoteBaseKey`"*" -and $error[0] -like "*`"The network path was not found.*")
        {
        $DisplayName = "No Network connection to $($Computer.Name)!"
        $DisplayVersion = "No Network connection to $($Computer.Name)!"
        $InstallLocation = "No Network connection to $($Computer.Name)!"
        $Publisher = "No Network connection to $($Computer.Name)!"
        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "No Network connection to $($Computer.Name)!"
        $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value "No Network connection to $($Computer.Name)!"
        $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value "No Network connection to $($Computer.Name)!"
        $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "No Network connection to $($Computer.Name)!"
        $subkeyarray += $obj
        $obj = $null
        }
        Else
        {
        #Drill down into the Uninstall key using the OpenSubKey Method
        $regkey = $reg.OpenSubKey($UninstallKey)
        #Retrieve an array of strings that contain all the subkey names
        $subkeys = $regkey.GetSubKeyNames() 
        #Open each Subkey and use GetValue Method to return the required values for each
            ForEach ($key in $subkeys)
            {
            $thisKey = $UninstallKey + "\\" + $key
            $thisSubKey = $reg.OpenSubKey($thisKey)
                If ($($thisSubKey.GetValue("DisplayName")) -like "*$Software*")
                {
                $obj = New-Object PSObject
                $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
                $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
                $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
                $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
                $subkeyarray += $obj
                $obj = $null
                }
            }
            If ($subkeyarray.DisplayName -notlike "*No Network connection*" -and $subkeyarray.DisplayName -notlike "*$Software*")
            {
            $obj = New-Object PSObject
            $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "$Software not installed!"
            $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value "$Software not installed!"
            $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value "$Software not installed!"
            $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "$Software not installed!"
            $subkeyarray += $obj
            $obj = $null
            }
        $DisplayName = [string]::Concat($subkeyarray.DisplayName)
        $DisplayVersion = [string]::Concat($subkeyarray.DisplayVersion)
        $InstallLocation = [string]::Concat($subkeyarray.InstallLocation)
        $Publisher = [string]::Concat($subkeyarray.Publisher)
        }
    $comp = New-Object PSObject
    $comp | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $Name
    $comp | Add-Member -MemberType NoteProperty -Name "IP_Address" -Value $IPAddress
    $comp | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS
    $comp | Add-Member -MemberType NoteProperty -Name "OUStructure" -Value $CanonicalName
    $comp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
    $comp | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $DisplayVersion
    $comp | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $InstallLocation
    $comp | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $Publisher
    $comp | Export-Csv -Path $OutputPath\$Software-Machines_$date.csv -Encoding ascii -Append -Force
    }
    Else
    {
    $nocomp = New-Object PSObject
    $nocomp | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value "$($Computer.Name) has not contacted AD in over $InactivityThreshold days"
    $nocomp | Add-Member -MemberType NoteProperty -Name "IP_Address" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "OUStructure" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "ERROR"
    $nocomp | Export-Csv -Path $OutputPath\$Software-Machines_$date.csv -Encoding ascii -Append -Force
    }
}
}
Function Remove-HashTableEmptyValue
{
<#
.SYNOPSIS
    This Function will remove the empty or Null entry of a hashtable object
.DESCRIPTION
    This Function will remove the empty or Null entry of a hashtable object
.PARAMETER Hashtable
    Specifies the hashtable that will be cleaned up.
.EXAMPLE
    Remove-HashTableEmptyValue -HashTable $SplattingVariable
.NOTES
    Francois-Xavier Cat
    @lazywinadm
    www.lazywinadmin.com
    github.com/lazywinadmin
#>
    [CmdletBinding()]
    PARAM([System.Collections.Hashtable]$HashTable)

    $HashTable.GetEnumerator().name |
        ForEach-Object -Process {
            if($HashTable[$_] -eq "" -or $HashTable[$_] -eq $null)
            {
                Write-Verbose -Message "[Remove-HashTableEmptyValue][PROCESS] - Property: $_ removing..."
                [void]$HashTable.Remove($_)
                Write-Verbose -Message "[Remove-HashTableEmptyValue][PROCESS] - Property: $_ removed"
            }
        }
}
#Requires -Version 5

<#
.SYNOPSIS
    A small wrapper for PowerShellGet to remove all older installed modules.
.DESCRIPTION
    A small wrapper for PowerShellGet to remove all older installed modules.
.PARAMETER ModuleName
    Name of a module to check and remove old versions of.
.EXAMPLE
    PS> Remove-OldModules.ps1

    Removes old modules installed via PowerShellGet.

.EXAMPLE
    PS> Remove-OldModules.ps1 -whatif

    Shows what old modules might be removed via PowerShellGet.

.NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 5.0

       Version History
       1.0.0 - Initial release
#>
[CmdletBinding( SupportsShouldProcess = $true )]
Param (
    [Parameter(HelpMessage = 'Name of a module to check and remove old versions of.')]
    [string]$ModuleName = '*'
)

try {
    Import-Module PowerShellGet
}
catch {
    Write-Warning 'Unable to load PowerShellGet. This script only works with PowerShell 5 and greater.'
    return
}
$WhatIfParam = @{}
$WhatIfParam.WhatIf = $WhatIf
Get-InstalledModule $ModuleName | Foreach-Object {
    $InstalledModules = get-module $_.Name -ListAvailable
    if ($InstalledModules.Count -gt 1) {
        $SortedModules = $InstalledModules | sort-object Version -Descending
        Write-Output "Multiple Module versions for the $($SortedModules[0].Name) module found. Highest version is: $($SortedModules[0].Version.ToString())"
        for ($index = 1; $index -lt $SortedModules.Count; $index++) {
            try {
                if ($pscmdlet.ShouldProcess( "$($SortedModules[$index].Name) - $($SortedModules[$index].Version)")) {
                    Write-Output "..Attempting to uninstall $($SortedModules[$index].Name) - Version $($SortedModules[$index].Version)"
                    Uninstall-Module -Name $SortedModules[$index].Name -MaximumVersion $SortedModules[$index].Version -ErrorAction Stop -Force
                }
            }
            catch {
                Write-Warning "Unable to remove module version $($SortedModules[$index].Version)"
            }
        }
    }
}
.\Upgrade-InstalledModules.ps1
Function Remove-PSObjectEmptyOrNullProperty
{
<#
	.SYNOPSIS
		Function to Remove all the empty or null properties with empty value in a PowerShell Object
	
	.DESCRIPTION
		Function to Remove all the empty or null properties with empty value in a PowerShell Object
	
	.PARAMETER PSObject
		Specifies the PowerShell Object
	
	.EXAMPLE
		PS C:\> Remove-PSObjectEmptyOrNullProperty -PSObject $UserInfo
	
	.NOTES
		Francois-Xavier Cat	
		www.lazywinadmin.com
		@lazywinadm
#>
	PARAM (
		$PSObject)
	PROCESS
	{
		$PsObject.psobject.Properties |
		Where-Object { -not $_.value } |
		ForEach-Object {
			$PsObject.psobject.Properties.Remove($_.name)
		}
	}
}
Function Remove-PSObjectProperty
{
<#
	.SYNOPSIS
		Function to Remove a specifid property from a PowerShell object
	
	.DESCRIPTION
		Function to Remove a specifid property from a PowerShell object
	
	.PARAMETER PSObject
		Specifies the PowerShell Object
	
	.PARAMETER Property
		Specifies the property to remove
	
	.EXAMPLE
		PS C:\> Remove-PSObjectProperty -PSObject $UserInfo -Property Info
	
	.NOTES
		Francois-Xavier Cat	
		www.lazywinadmin.com
		@lazywinadm
#>
	PARAM (
		$PSObject,
		
		[String[]]$Property)
	PROCESS
	{
		Foreach ($item in $Property)
		{
			$PSObject.psobject.Properties.Remove("$item")
		}
	}
}
Function Remove-StringDiacritic
{
<#
.SYNOPSIS
	This Function will remove the diacritics (accents) characters from a string.
	
.DESCRIPTION
	This Function will remove the diacritics (accents) characters from a string.

.PARAMETER String
	Specifies the String(s) on which the diacritics need to be removed

.PARAMETER NormalizationForm
	Specifies the normalization form to use
	https://msdn.microsoft.com/en-us/library/system.text.normalizationform(v=vs.110).aspx

.EXAMPLE
	PS C:\> Remove-StringDiacritic "L'été de Raphaël"
	
	L'ete de Raphael

.NOTES
	Francois-Xavier Cat
	@lazywinadm
	www.lazywinadmin.com
	github.com/lazywinadmin
#>
	[CMdletBinding()]
	PARAM
	(
		[ValidateNotNullOrEmpty()]
		[Alias('Text')]
		[System.String[]]$String,
		[System.Text.NormalizationForm]$NormalizationForm = "FormD"
	)
	
	FOREACH ($StringValue in $String)
	{
		Write-Verbose -Message "$StringValue"
		try
		{	
			# Normalize the String
			$Normalized = $StringValue.Normalize($NormalizationForm)
			$NewString = New-Object -TypeName System.Text.StringBuilder
			
			# Convert the String to CharArray
			$normalized.ToCharArray() |
			ForEach-Object -Process {
				if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -ne [Globalization.UnicodeCategory]::NonSpacingMark)
				{
					[void]$NewString.Append($psitem)
				}
			}

			#Combine the new string chars
			Write-Output $($NewString -as [string])
		}
		Catch
		{
			Write-Error -Message $Error[0].Exception.Message
		}
	}
}
Function Remove-StringLatinCharacter
{
<#
.SYNOPSIS
    Function to remove diacritics from a string
.PARAMETER String
	Specifies the String that will be processed
.EXAMPLE
    Remove-StringLatinCharacter -String "L'été de Raphaël"

    L'ete de Raphael
.EXAMPLE
    Foreach ($file in (Get-ChildItem c:\test\*.txt))
    {
        # Get the content of the current file and remove the diacritics
        $NewContent = Get-content $file | Remove-StringLatinCharacter
    
        # Overwrite the current file with the new content
        $NewContent | Set-Content $file
    }

    Remove diacritics from multiple files

.NOTES
    Francois-Xavier Cat
    lazywinadmin.com
    @lazywinadm
    github.com/lazywinadmin

    BLOG ARTICLE
        http://www.lazywinadmin.com/2015/05/powershell-remove-diacritics-accents.html
	
    VERSION HISTORY
        1.0.0.0 | Francois-Xavier Cat
            Initial version Based on Marcin Krzanowic code
        1.0.0.1 | Francois-Xavier Cat
            Added support for ValueFromPipeline
        1.0.0.2 | Francois-Xavier Cat
            Add Support for multiple String
            Add Error Handling
#>
    [CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline=$true)]
		[System.String[]]$String
		)
	PROCESS
	{
        FOREACH ($StringValue in $String)
        {
            Write-Verbose -Message "$StringValue"

            TRY
            {
                [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($StringValue))
            }
		    CATCH
            {
                Write-Error -Message $Error[0].exception.message
            }
        }
	}
}
Function Remove-StringSpecialCharacter
{
<#
.SYNOPSIS
	This Function will remove the special character from a string.
	
.DESCRIPTION
	This Function will remove the special character from a string.
	I'm using Unicode Regular Expressions with the following categories
	\p{L} : any kind of letter from any language.
	\p{Nd} : a digit zero through nine in any script except ideographic 
	
	http://www.regular-expressions.info/unicode.html
	http://unicode.org/reports/tr18/

.PARAMETER String
	Specifies the String on which the special character will be removed

.SpecialCharacterToKeep
	Specifies the special character to keep in the output

.EXAMPLE
	PS C:\> Remove-StringSpecialCharacter -String "^&*@wow*(&(*&@"
	wow
.EXAMPLE
	PS C:\> Remove-StringSpecialCharacter -String "wow#@!`~)(\|?/}{-_=+*"
	
	wow
.EXAMPLE
	PS C:\> Remove-StringSpecialCharacter -String "wow#@!`~)(\|?/}{-_=+*" -SpecialCharacterToKeep "*","_","-"
	wow-_*

.NOTES
	Francois-Xavier Cat
	@lazywinadm
	www.lazywinadmin.com
	github.com/lazywinadmin
#>
	[CmdletBinding()]
	param
	(
		[Parameter(ValueFromPipeline)]
		[ValidateNotNullOrEmpty()]
		[Alias('Text')]
		[System.String[]]$String,
		
		[Alias("Keep")]
		[ValidateNotNullOrEmpty()]
		[String[]]$SpecialCharacterToKeep
	)
	PROCESS
	{
		IF ($PSBoundParameters["SpecialCharacterToKeep"])
		{
			$Regex = "[^\p{L}\p{Nd}"
			Foreach ($Character in $SpecialCharacterToKeep)
			{
				$Regex += "/$character"
			}
			
			$Regex += "]+"
		} #IF($PSBoundParameters["SpecialCharacterToKeep"])
		ELSE { $Regex = "[^\p{L}\p{Nd}]+" }
		
		FOREACH ($Str in $string)
		{
			Write-Verbose -Message "Original String: $Str"
			$Str -replace $regex, ""
		}
	} #PROCESS
}
#PowerShell Script Containing Function Used to Remove User Profiles & Additional Remnants of C:\Users Directory
#Developer: Andrew Saraceni (saraceni@wharton.upenn.edu)
#Date: 12/22/14

#Requires -Version 2.0

Function Remove-UserProfile
{
    <#
    .SYNOPSIS
    Removes user profiles and additional contents of the C:\Users 
    directory if specified.
    .DESCRIPTION
    Gathers a list of profiles to be removed from the local computer, 
    passing on exceptions noted via the Exclude parameter and/or 
    profiles newer than the date specified via the Before parameter.  
    If desired, additional files and folders within C:\Users can also 
    be removed via use of the DirectoryCleanup parameter.

    Once gathered, miscellaneous items are first removed from the 
    C:\Users directory if specified, followed by the profile objects 
    themselves and all associated registry keys per profile.  A listing 
    of current items within the C:\Users directory is returned 
    following the profile removal process.
    .PARAMETER Exclude
    Specifies one or more profile names to exclude from the removal 
    process.
    .PARAMETER Before
    Specifies a date from which to remove profiles before that haven't 
    been accessed since that date.
    .PARAMETER DirectoryCleanup
    Removes additional files/folders (i.e. non-profiles) within the 
    C:\Users directory.
    .EXAMPLE
    Remove-UserProfile
    Remove all non-active and non-system designated user profiles 
    from the local computer.
    .EXAMPLE
    Remove-UserProfile -Before (Get-Date).AddMonths(-1) -Verbose
    Remove all non-active and non-system designated user profiles 
    not used within the past month, displaying verbose output as well.
    .EXAMPLE
    Remove-UserProfile -Exclude @("labadmin", "desktopuser") -DirectoryCleanup
    Remove all non-active and non-system designated user profiles 
    except "labadmin" and "desktopuser", and remove additional 
    non-profile files/folders within C:\Users as well.
    .NOTES
    Even when not specifying the Exclude parameter, the following 
    profiles are not removed when utilizing this cmdlet:
    C:\Windows\ServiceProfiles\NetworkService 
    C:\Windows\ServiceProfiles\LocalService 
    C:\Windows\system32\config\systemprofile 
    C:\Users\Public
    C:\Users\Default

    Aside from the original profile directory (within C:\Users) 
    itself, the following registry items are also cleared upon 
    profile removal via WMI:
    "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{SID of User}"
    "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileGuid\{GUID}" SidString = {SID of User}
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\{SID of User}"

    Additionally, any currently loaded/in use profiles will not be 
    removed.  Regarding miscellaneous non-profile items, hidden items 
    are not enumerated or removed from C:\Users during this process.

    This cmdlet requires adminisrative privileges to run effectively.
      
    This cmdlet is not intended to be used on Virtual Desktop 
    Infrastructure (VDI) environments or others which utilize 
    persistent storage on alternate disks, or any configurations 
    which utilize another directory other than C:\Users to store 
    user profiles.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$false)]
        [String[]]$Exclude,
        [Parameter(Position=1,Mandatory=$false)]
        [DateTime]$Before,
        [Parameter(Position=2,Mandatory=$false)]
        [Switch]$DirectoryCleanup
    )

    Write-Verbose "Gathering List of Profiles on $env:COMPUTERNAME to Remove..."

    $userProfileFilter = "Loaded = 'False' AND Special = 'False'"
    $cleanupExclusions = @("Public", "Default")

    if ($Exclude)
    {
        foreach ($exclusion in $Exclude)
        {
            $userProfileFilter += "AND NOT LocalPath LIKE '%$exclusion'"
            $cleanupExclusions += $exclusion
        }
    }

    if ($Before)
    {
        $userProfileFilter += "AND LastUseTime < '$Before'"

        $keepUserProfileFilter = "Special = 'False' AND LastUseTime >= '$Before'"
        $profilesToKeep = Get-WmiObject -Class Win32_UserProfile -Filter $keepUserProfileFilter -ErrorAction Stop

        foreach ($profileToKeep in $profilesToKeep)
        {
            try
            {
                $userSID = New-Object -TypeName System.Security.Principal.SecurityIdentifier($($profileToKeep.SID))
                $userName = $userSID.Translate([System.Security.Principal.NTAccount])
                
                $keepUserName = $userName.Value -replace ".*\\", ""
                $cleanupExclusions += $keepUserName
            }
            catch [System.Security.Principal.IdentityNotMappedException]
            {
                Write-Warning "Cannot Translate SID to UserName - Not Adding Value to Exceptions List"
            }
        }
    }

    $profilesToDelete = Get-WmiObject -Class Win32_UserProfile -Filter $userProfileFilter -ErrorAction Stop

    if ($DirectoryCleanup)
    {
        $usersChildItem = Get-ChildItem -Path "C:\Users" -Exclude $cleanupExclusions

        foreach ($usersChild in $usersChildItem)
        {
            if ($profilesToDelete.LocalPath -notcontains $usersChild.FullName)
            {    
                try
                {
                    Write-Verbose "Additional Directory Cleanup - Removing $($usersChild.Name) on $env:COMPUTERNAME..."
                    
                    Remove-Item -Path $($usersChild.FullName) -Recurse -Force -ErrorAction Stop
                }
                catch [System.InvalidOperationException]
                {
                    Write-Verbose "Skipping Removal of $($usersChild.Name) on $env:COMPUTERNAME as Item is Currently In Use..."
                }
            }
        }
    }

    foreach ($profileToDelete in $profilesToDelete)
    {
        Write-Verbose "Removing Profile $($profileToDelete.LocalPath) & Associated Registry Keys on $env:COMPUTERNAME..."
                
        Remove-WmiObject -InputObject $profileToDelete -ErrorAction Stop
    }

    $finalChildItem = Get-ChildItem -Path "C:\Users" | Select-Object -Property Name, FullName, LastWriteTime
                
    return $finalChildItem
}
Function Remove-VMEvcMode {
<#  
.SYNOPSIS  
    Removes the EVC status of a VM
.DESCRIPTION 
    Will remove the EVC status for the specified VM
.NOTES  
    Author:  Kyle Ruddy, @kmruddy, thatcouldbeaproblem.com
.PARAMETER Name
    VM name which the Function should be ran against
.EXAMPLE
	Remove-VMEvcMode -Name vmName
	Removes the EVC status of the provided VM 
#>
[CmdletBinding()] 
	param(
	[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        $Name
  	)

    Process {
        $evVM = @()
        $updateVM = @()

        if ($name -is [string]) {$evVM += Get-VM -Name $Name -ErrorAction SilentlyContinue}
        elseif ($name -is [array]) {

            if ($name[0] -is [string]) {
                $name | foreach {
                    $evVM += Get-VM -Name $_ -ErrorAction SilentlyContinue
                }
            }
            elseif ($name[0] -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM = $name}

        }
        elseif ($name -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM += $name}
        
        if ($evVM -eq $null) {Write-Warning "No VMs found."}
        else {
            foreach ($v in $evVM) {

                if (($v.HardwareVersion -ge 'vmx-14' -and $v.PowerState -eq 'PoweredOff') -or ($v.Version -ge 'v14' -and $v.PowerState -eq 'PoweredOff')) {

                    $v.ExtensionData.ApplyEvcModeVM_Task($null, $true) | Out-Null
                    $updateVM += $v.Name
                                        
                }
                else {Write-Warning $v.Name + " does not have the minimum requirements of being Hardware Version 14 and powered off."}

            }

            if ($updateVM) {
            
            Start-Sleep -Seconds 2
            Get-VMEvcMode -Name $updateVM
            
            }

        }

    }

}
Function Resolve-ShortURL
{
<#
.SYNOPSIS
	Function to resolve a short URL to the absolute URI

.DESCRIPTION
	Function to resolve a short URL to the absolute URI

.PARAMETER ShortUrl
	Specifies the ShortURL

.EXAMPLE
	Resolve-ShortURL -ShortUrl http://goo.gl/P5PKq

.EXAMPLE
	Resolve-ShortURL -ShortUrl http://go.microsoft.com/fwlink/?LinkId=280243

.NOTES
	Francois-Xavier Cat
	lazywinadmin.com
	@lazywinadm
	github.com/lazywinadmin
#>
	
	[CmdletBinding()]
	[OutputType([System.String])]
	PARAM
	(
		[String[]]$ShortUrl
	)
	
	FOREACH ($URL in $ShortUrl)
	{
		TRY
		{
			Write-Verbose -Message "$URL - Querying..."
			(Invoke-WebRequest -Uri $URL -MaximumRedirection 0 -ErrorAction Ignore).Headers.Location
		}
		CATCH
		{
			Write-Error -Message $Error[0].Exception.Message
		}
	}
}
Function Restart-PowershellHost
{
    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact='High')] 
    Param
    (
        [switch]
        $AsAdministrator,

        [switch]
        $Force
    )
    
    process
    {
        if ($Force -or $PSCmdlet.ShouldProcess($proc.Name, "Restart the console as administrator : '{0}'" -f $AsAdministrator))    # comfirmation to restart
        {
            if (($host.Name -eq 'Windows PowerShell ISE Host') -and ($psISE.PowerShellTabs.Files.IsSaved -contains $false))        # ise detect and unsave tab check
            {
                if ($Force -or $PSCmdlet.ShouldProcess('Unsaved work detected?','Unsaved work detected. Save changes?','Confirm')) # ise tab save dialog
                {
                    # dialog selected yes.
                    $psISE.PowerShellTabs | Start-SaveAndCloseISETabs
                }
                else
                {
                    # dialog selected no.
                    $psISE.PowerShellTabs | Start-CloseISETabs
                }
            }

            #region restart host process
            Write-Debug ("Start new host : '{0}'" -f $proc.Name)
            Start-Process @params

            Write-Debug ("Close old host : '{0}'" -f $proc.Name)
            $proc.CloseMainWindow()
            #endregion
        }
    }

    begin
    {
        $proc = Get-Process -Id $PID
 
        #region Setup parameter for restart host
        $params = @{
            FilePath = $proc.Path
        }

        if ($AsAdministrator)
        {
            $params.Verb = 'runas'
        }

        if ($cmdArgs)
        {
            $params.ArgumentList = [Environment]::GetCommandLineArgs() | Select-Object -Skip 1
        }
        #endregion

        #region internal Function to close ise with save
        filter Start-SaveAndCloseISETabs
        {
            $_.Files `
            | % { 
                if($_.IsUntitled -and (-not $_.IsSaved))
                {
                    $_.SaveAs($_.FullPath, [System.Text.Encoding]::UTF8)
                }
                elseif(-not $_.IsSaved)
                {
                    $_.Save()
                }
            }
        }
        #endregion

        #region internal Function to close ise without save
        filter Start-CloseISETabs
        {
            $ISETab = $_
            $unsavedFiles = $IseTab.Files | where IsSaved -eq $false
            $unsavedFiles | % {$IseTab.Files.Remove($_,$true)}
        }
        #endregion
    }
}
#  This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
#  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR   
#  FITNESS FOR A PARTICULAR PURPOSE.
#
#
#  Script queries port 135 to get the listening ephemeral ports from the remote server
#  and verifies that they are reachable.  
#
#
#  Usage:  RPCCheck -Server YourServerNameHere
#
#
#  Note:  The script relies on portqry.exe (from Sysinternals) to get port 135 output.
#  The path to portqry.exe will need to be modified to reflect your location
#  


Param(
  [string]$Server
)

#  WORKFLOW QUERIES THE PASSED ARRAY OF PORTS TO DETERMINE STATUS

workflow Check-Port {

  param ([string[]]$RPCServer,[array]$arrRPCPorts)

  $comp = hostname

 

  ForEach -parallel ($RPCPort in $arrRPCPorts)
  {

      $bolResult = InlineScript{Test-NetConnection -ComputerName $Using:RPCServer -port $Using:RPCPort -InformationLevel Quiet}

 If ($bolResult)
 {
      Write-Output "$RPCPort on $RPCServer is reachable"
 }
 Else
 {
     Write-Output "$RPCPort on $RPCServer is unreachable"
 }
}
}

#  INITIAL RPC PORT

$strRPCPort = "135"

#  MODIFY PATH TO THE PORTQRY BINARY IF NECESSARY
$strPortQryPath = "C:\Sysinternals"

#  TEST THE PATH TO SEE IF THE BINARY EXISTS
If (Test-Path "$strPortQryPath\PortQry.exe")
{
  $strPortQryCmd = "$strPortQryPath\PortQry.exe -e $strRPCPort -n $Server"
}
Else
{
  Write-Output "Could not locate Portqry.exe at the path $strPortQryPath"
  Exit
}

#  CREATE AN EMPTY ARRAY TO HOLD THE PORTS RETURNED FROM THE RPC PORTMAPPER
$arrPorts = @()

#  RUN THE PORTQRY COMMAND TO GET THE EPHEMERAL PORTS

$arrQuryResult = Invoke-Expression $strPortQryCmd

# CREATE AN ARRAY OF THE PORTS
ForEach ($strResult in $arrQuryResult)
{
  If ($strResult.Contains("ip_tcp"))
  {
  $arrSplt = $strResult.Split("[")
  $strPort = $arrSplt[1]
  $strPort = $strPort.Replace("]","")
  $arrPorts += $strPort
  }
}

#  DE-DUPLICATE THE PORTS
$arrPorts = $arrPorts | Sort-Object |Select-Object -Unique

#  EXECUTE THE WORKFLOW TO CHECK THE PORTS
Check-Port -RPCServer $Server -arrRPCPorts $arrPorts
Function Set-ScreenResolution { 
 
<# 
    .Synopsis 
        Sets the Screen Resolution of the primary monitor 
    .Description 
        Uses Pinvoke and ChangeDisplaySettings Win32API to make the change 
    .Example 
        Set-ScreenResolution -Width 1024 -Height 768         
    #> 
param ( 
[Parameter(Mandatory=$true, 
           Position = 0)] 
[int] 
$Width, 
 
[Parameter(Mandatory=$true, 
           Position = 1)] 
[int] 
$Height 
) 
 
$pinvokeCode = @" 
 
using System; 
using System.Runtime.InteropServices; 
 
namespace Resolution 
{ 
 
    [StructLayout(LayoutKind.Sequential)] 
    public struct DEVMODE1 
    { 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmDeviceName; 
        public short dmSpecVersion; 
        public short dmDriverVersion; 
        public short dmSize; 
        public short dmDriverExtra; 
        public int dmFields; 
 
        public short dmOrientation; 
        public short dmPaperSize; 
        public short dmPaperLength; 
        public short dmPaperWidth; 
 
        public short dmScale; 
        public short dmCopies; 
        public short dmDefaultSource; 
        public short dmPrintQuality; 
        public short dmColor; 
        public short dmDuplex; 
        public short dmYResolution; 
        public short dmTTOption; 
        public short dmCollate; 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmFormName; 
        public short dmLogPixels; 
        public short dmBitsPerPel; 
        public int dmPelsWidth; 
        public int dmPelsHeight; 
 
        public int dmDisplayFlags; 
        public int dmDisplayFrequency; 
 
        public int dmICMMethod; 
        public int dmICMIntent; 
        public int dmMediaType; 
        public int dmDitherType; 
        public int dmReserved1; 
        public int dmReserved2; 
 
        public int dmPanningWidth; 
        public int dmPanningHeight; 
    }; 
 
 
 
    class User_32 
    { 
        [DllImport("user32.dll")] 
        public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE1 devMode); 
        [DllImport("user32.dll")] 
        public static extern int ChangeDisplaySettings(ref DEVMODE1 devMode, int flags); 
 
        public const int ENUM_CURRENT_SETTINGS = -1; 
        public const int CDS_UPDATEREGISTRY = 0x01; 
        public const int CDS_TEST = 0x02; 
        public const int DISP_CHANGE_SUCCESSFUL = 0; 
        public const int DISP_CHANGE_RESTART = 1; 
        public const int DISP_CHANGE_FAILED = -1; 
    } 
 
 
 
    public class PrmaryScreenResolution 
    { 
        static public string ChangeResolution(int width, int height) 
        { 
 
            DEVMODE1 dm = GetDevMode1(); 
 
            if (0 != User_32.EnumDisplaySettings(null, User_32.ENUM_CURRENT_SETTINGS, ref dm)) 
            { 
 
                dm.dmPelsWidth = width; 
                dm.dmPelsHeight = height; 
 
                int iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_TEST); 
 
                if (iRet == User_32.DISP_CHANGE_FAILED) 
                { 
                    return "Unable To Process Your Request. Sorry For This Inconvenience."; 
                } 
                else 
                { 
                    iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_UPDATEREGISTRY); 
                    switch (iRet) 
                    { 
                        case User_32.DISP_CHANGE_SUCCESSFUL: 
                            { 
                                return "Success"; 
                            } 
                        case User_32.DISP_CHANGE_RESTART: 
                            { 
                                return "You Need To Reboot For The Change To Happen.\n If You Feel Any Problem After Rebooting Your Machine\nThen Try To Change Resolution In Safe Mode."; 
                            } 
                        default: 
                            { 
                                return "Failed To Change The Resolution"; 
                            } 
                    } 
 
                } 
 
 
            } 
            else 
            { 
                return "Failed To Change The Resolution."; 
            } 
        } 
 
        private static DEVMODE1 GetDevMode1() 
        { 
            DEVMODE1 dm = new DEVMODE1(); 
            dm.dmDeviceName = new String(new char[32]); 
            dm.dmFormName = new String(new char[32]); 
            dm.dmSize = (short)Marshal.SizeOf(dm); 
            return dm; 
        } 
    } 
} 
 
"@ 
 
Add-Type $pinvokeCode -ErrorAction SilentlyContinue 
[Resolution.PrmaryScreenResolution]::ChangeResolution($width,$height) 
} 
Set-ScreenResolution -Width 1920 -Height 1080
Function Search-Registry {
<#
.SYNOPSIS
Searches registry key names, value names, and value data (limited).

.DESCRIPTION
This Function can search registry key names, value names, and value data (in a limited fashion). It outputs custom objects that contain the key and the first match type (KeyName, ValueName, or ValueData).

.EXAMPLE
Search-Registry -Path HKLM:\SYSTEM\CurrentControlSet\Services\* -SearchRegex "svchost" -ValueData

.EXAMPLE
Search-Registry -Path HKLM:\SOFTWARE\Microsoft -Recurse -ValueNameRegex "ValueName1|ValueName2" -ValueDataRegex "ValueData" -KeyNameRegex "KeyNameToFind1|KeyNameToFind2"

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0, ValueFromPipelineByPropertyName)]
        [Alias("PsPath")]
        # Registry path to search
        [string[]] $Path,
        # Specifies whether or not all subkeys should also be searched
        [switch] $Recurse,
        [Parameter(ParameterSetName="SingleSearchString", Mandatory)]
        # A regular expression that will be checked against key names, value names, and value data (depending on the specified switches)
        [string] $SearchRegex,
        [Parameter(ParameterSetName="SingleSearchString")]
        # When the -SearchRegex parameter is used, this switch means that key names will be tested (if none of the three switches are used, keys will be tested)
        [switch] $KeyName,
        [Parameter(ParameterSetName="SingleSearchString")]
        # When the -SearchRegex parameter is used, this switch means that the value names will be tested (if none of the three switches are used, value names will be tested)
        [switch] $ValueName,
        [Parameter(ParameterSetName="SingleSearchString")]
        # When the -SearchRegex parameter is used, this switch means that the value data will be tested (if none of the three switches are used, value data will be tested)
        [switch] $ValueData,
        [Parameter(ParameterSetName="MultipleSearchStrings")]
        # Specifies a regex that will be checked against key names only
        [string] $KeyNameRegex,
        [Parameter(ParameterSetName="MultipleSearchStrings")]
        # Specifies a regex that will be checked against value names only
        [string] $ValueNameRegex,
        [Parameter(ParameterSetName="MultipleSearchStrings")]
        # Specifies a regex that will be checked against value data only
        [string] $ValueDataRegex
    )

    begin {
        switch ($PSCmdlet.ParameterSetName) {
            SingleSearchString {
                $NoSwitchesSpecified = -not ($PSBoundParameters.ContainsKey("KeyName") -or $PSBoundParameters.ContainsKey("ValueName") -or $PSBoundParameters.ContainsKey("ValueData"))
                if ($KeyName -or $NoSwitchesSpecified) { $KeyNameRegex = $SearchRegex }
                if ($ValueName -or $NoSwitchesSpecified) { $ValueNameRegex = $SearchRegex }
                if ($ValueData -or $NoSwitchesSpecified) { $ValueDataRegex = $SearchRegex }
            }
            MultipleSearchStrings {
                # No extra work needed
            }
        }
    }

    process {
        foreach ($CurrentPath in $Path) {
            Get-ChildItem $CurrentPath -Recurse:$Recurse | 
                ForEach-Object {
                    $Key = $_

                    if ($KeyNameRegex) { 
                        Write-Verbose ("{0}: Checking KeyNamesRegex" -f $Key.Name) 
        
                        if ($Key.PSChildName -match $KeyNameRegex) { 
                            Write-Verbose "  -> Match found!"
                            return [PSCustomObject] @{
                                Key = $Key
                                Reason = "KeyName"
                            }
                        } 
                    }
        
                    if ($ValueNameRegex) { 
                        Write-Verbose ("{0}: Checking ValueNamesRegex" -f $Key.Name)
            
                        if ($Key.GetValueNames() -match $ValueNameRegex) { 
                            Write-Verbose "  -> Match found!"
                            return [PSCustomObject] @{
                                Key = $Key
                                Reason = "ValueName"
                            }
                        } 
                    }
        
                    if ($ValueDataRegex) { 
                        Write-Verbose ("{0}: Checking ValueDataRegex" -f $Key.Name)
            
                        if (($Key.GetValueNames() | % { $Key.GetValue($_) }) -match $ValueDataRegex) { 
                            Write-Verbose "  -> Match!"
                            return [PSCustomObject] @{
                                Key = $Key
                                Reason = "ValueData"
                            }
                        }
                    }
                }
        }
    }
}
Function Set-FileTime{
  param(
    [string[]]$paths,
    [bool]$only_modification = $false,
    [bool]$only_access = $false
  );

  begin {
    Function updateFileSystemInfo([System.IO.FileSystemInfo]$fsInfo) {
      $datetime = get-date
      if ( $only_access )
      {
         $fsInfo.LastAccessTime = $datetime
      }
      elseif ( $only_modification )
      {
         $fsInfo.LastWriteTime = $datetime
      }
      else
      {
         $fsInfo.CreationTime = $datetime
         $fsInfo.LastWriteTime = $datetime
         $fsInfo.LastAccessTime = $datetime
       }
    }
   
    Function touchExistingFile($arg) {
      if ($arg -is [System.IO.FileSystemInfo]) {
        updateFileSystemInfo($arg)
      }
      else {
        $resolvedPaths = resolve-path $arg
        foreach ($rpath in $resolvedPaths) {
          if (test-path -type Container $rpath) {
            $fsInfo = new-object System.IO.DirectoryInfo($rpath)
          }
          else {
            $fsInfo = new-object System.IO.FileInfo($rpath)
          }
          updateFileSystemInfo($fsInfo)
        }
      }
    }
   
    Function touchNewFile([string]$path) {
      #$null > $path
      Set-Content -Path $path -value $null;
    }
  }
 
  process {
    if ($_) {
      if (test-path $_) {
        touchExistingFile($_)
      }
      else {
        touchNewFile($_)
      }
    }
  }
 
  end {
    if ($paths) {
      foreach ($path in $paths) {
        if (test-path $path) {
          touchExistingFile($path)
        }
        else {
          touchNewFile($path)
        }
      }
    }
  }
}
Function Set-GPOStatus {
<# comment based help is here #>

[cmdletbinding(SupportsShouldProcess)]

Param(
[Parameter(Position=0,Mandatory=$True,HelpMessage="Enter the name of a GPO",
ValueFromPipeline,ValueFromPipelinebyPropertyName)]
[Alias("name")]
[ValidateNotNullorEmpty()]
[Parameter(ParameterSetName="EnableAll")]
[Parameter(ParameterSetName="DisableAll")]
[Parameter(ParameterSetName="DisableUser")]
[Parameter(ParameterSetName="DisableComputer")]
[object]$DisplayName,
[Parameter(ParameterSetName="EnableAll")]
[Parameter(ParameterSetName="DisableAll")]
[Parameter(ParameterSetName="DisableUser")]
[Parameter(ParameterSetName="DisableComputer")]
[string]$Domain,
[Parameter(ParameterSetName="EnableAll")]
[Parameter(ParameterSetName="DisableAll")]
[Parameter(ParameterSetName="DisableUser")]
[Parameter(ParameterSetName="DisableComputer")]
[string]$Server,
[Parameter(ParameterSetName="EnableAll")]
[switch]$EnableAll,
[Parameter(ParameterSetName="DisableAll")]
[switch]$DisableAll,
[Parameter(ParameterSetName="DisableUser")]
[switch]$DisableUser,
[Parameter(ParameterSetName="DisableComputer")]
[switch]$DisableComputer,
[Parameter(ParameterSetName="EnableAll")]
[Parameter(ParameterSetName="DisableAll")]
[Parameter(ParameterSetName="DisableUser")]
[Parameter(ParameterSetName="DisableComputer")]
[switch]$Passthru
)

Begin {
    Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"  
       
    #define a hashtable we can for splatting
    $paramhash=@{ErrorAction="Stop"}
    if ($domain) { $paramhash.Add("Domain",$Domain) }
    if ($server) { $paramhash.Add("Server",$Server) }

} #begin

Process {
    #define appropriate GPO setting value depending on parameter
    Switch ($PSCmdlet.ParameterSetName) {
    "EnableAll" { $status = "AllSettingsEnabled" }
    "DisableAll" { $status = "AllSettingsDisabled" }
    "DisableUser" { $status = "UserSettingsEnabled" }
    "DisableComputer" { $status = "ComputerSettingsEnabled" }
    default {
            Write-Warning "You didn't specify a GPO setting. No changes will be made."
            Return
            }
    }
    
    #if GPO is a string, get it with Get-GPO
    if ($Displayname -is [string]) {
        $paramhash.Add("name",$DisplayName)
        
        Write-Verbose "Retrieving Group Policy Object"
        Try {
            write-verbose "Using Parameter hash $($paramhash | out-string)"
            $gpo=Get-GPO @paramhash
        }
        Catch {
            Write-Warning "Failed to find a GPO called $displayname"
            Return
        }
    }
    else {
        $paramhash.Add("GUID",$DisplayName.id)
        $gpo = $DisplayName
    }

    #set the GPOStatus property on the GPO object to the correct value. The change is immediate.
    Write-Verbose "Setting GPO $($gpo.displayname) status to $status"

    if ($PSCmdlet.ShouldProcess("$($gpo.Displayname) : $status ")) {
        $gpo.gpostatus=$status
        if ($passthru) {
            #refresh the GPO Object
            write-verbose "Using Parameter hash $($paramhash | out-string)"
            get-gpo @paramhash 
        }
    } #should process

} #process

End {
    Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
} #end

} #end Set-GPOStatus
Function Set-NetworkLevelAuthentication
{
<#
.SYNOPSIS
	This Function will set the NLA setting on a local machine or remote machine

.DESCRIPTION
	This Function will set the NLA setting on a local machine or remote machine

.PARAMETER  ComputerName
	Specify one or more computers

.PARAMETER EnableNLA
	Specify if the NetworkLevelAuthentication need to be set to $true or $false

.PARAMETER  Credential
	Specify the alternative credential to use. By default it will use the current one.

.EXAMPLE
	Set-NetworkLevelAuthentication -EnableNLA $true

.EXAMPLE
	Set-NetworkLevelAuthentication -EnableNLA $true -computername "SERVER01","SERVER02"

.EXAMPLE
	Set-NetworkLevelAuthentication -EnableNLA $true -computername (Get-Content ServersList.txt)

.NOTES
	DATE	: 2014/04/01
	AUTHOR	: Francois-Xavier Cat
	WWW		: http://lazywinadmin.com
	Twitter	: @lazywinadm

	Article : http://lazywinadmin.com/2014/04/powershell-getset-network-level.html
	GitHub	: https://github.com/lazywinadmin/PowerShell
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[System.String[]]$ComputerName = $env:ComputerName,
		
		[Parameter(Mandatory)]
		[System.Boolean]$EnableNLA,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message '[BEGIN] Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "[BEGIN] Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			Write-Verbose -message $Computer
			TRY
			{
				# Building Splatting for CIM Sessions
				Write-Verbose -message "$Computer - CIM/WIM - Building Splatting"
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the Function
				IF ($PSBoundParameters['Credential'])
				{
					Write-Verbose -message "[PROCESS] $Computer - CIM/WMI - Add Credential Specified"
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "[PROCESS] $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -Count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "[PROCESS] $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "[PROCESS] $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "[PROCESS] $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "[PROCESS] $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "[PROCESS] $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				$NLAinfo | Invoke-CimMethod -MethodName SetUserAuthenticationRequired -Arguments @{ UserAuthenticationRequired = $EnableNLA } -ErrorAction 'Continue' -ErrorVariable ErrorProcessInvokeWmiMethod
			}
			
			CATCH
			{
				Write-Warning -Message "Error on $Computer"
				Write-Error -Message $_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "[PROCESS] Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "[PROCESS] Error - $ProcessError" }
				if ($ErrorProcessInvokeWmiMethod) { Write-Warning -Message "[PROCESS] Error - $ErrorProcessInvokeWmiMethod" }
			}#CATCH
			FINALLY
			{
				if ($CimSession)
				{
					# CLeanup/Close the remaining session
					Write-Verbose -Message "[PROCESS] Finally Close any CIM Session(s)"
					Remove-CimSession -CimSession $CimSession
				}
			}
		} # FOREACH
	}#PROCESS
	END
	{
		Write-Verbose -Message "[END] Script is completed"
	}
}
Function Set-PowerShellMemoryTuning{

    param(
        [parameter(
            position = 0,
            mandatory = 1)]
        [ValidateNotNullorEmpty()]
        [ValidateRange(1,2147483647)]
        [int]
        $memory
    )

    # Test Elevated or not
    $TestElevated = {
        $user = [Security.Principal.WindowsIdentity]::GetCurrent()
        (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

    if (&$TestElevated)
    {

        # Machine Wide Memory Tuning
        Write-Warning "Current Memory for Machine wide is : $((Get-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB).value) MB"

        Write-Warning "Change Memory for Machine wide to : $memory MB"
        Set-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB $memory


        # EndPoing Memory Tuning
        Write-Warning "Current Memory for Plugin is : $((Get-Item WSMan:localhost\Plugin\microsoft.powershell\Quotas\MaxConcurrentCommandsPerShell).value) MB"

        Write-Warning "Change Memory for Plugin to : $memory MB"
        Set-Item WSMan:localhost\Plugin\microsoft.powershell\Quotas\MaxConcurrentCommandsPerShell $memory


        # Restart WinRM
        Write-Warning "Restarting WinRM"
        Restart-Service WinRM -Force -PassThru


        # Show Current parameters
        Write-Warning "Current Memory for Machine wide is : $((Get-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB).value) MB"
        Write-Warning "Current Memory for Plugin is : $((Get-Item WSMan:localhost\Plugin\microsoft.powershell\Quotas\MaxConcurrentCommandsPerShell).value) MB"
    }
    else
    {
        Write-Error "This Cmdlet requires Admin right. Please Elevate and try again."
    }

}
Function Set-PowerShellWindowTitle
{
<#
	.SYNOPSIS
		Function to set the title of the PowerShell Window
	
	.DESCRIPTION
		Function to set the title of the PowerShell Window
	
	.PARAMETER Title
		Specifies the Title of the PowerShell Window
	
	.EXAMPLE
		PS C:\> Set-PowerShellWindowTitle -Title LazyWinAdmin.com
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	PARAM($Title)
	$Host.UI.RawUI.WindowTitle = $Title	
}

Function Enable-PSScriptBlockLogging 
{  
    $basePath = "HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging"  
 
    if(-not (Test-Path $basePath))  
    {  
        $null = New-Item $basePath –Force  
    }
   
    Set-ItemProperty $basePath -Name EnableScriptBlockLogging -Value "1" 
}
 

Function Disable-PSScriptBlockLogging 
{  
    Remove-Item HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging -Force –Recurse 
}
 

Function Enable-PSScriptBlockInvocationLogging 
{
    $basePath = "HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging"  
 
    if(-not (Test-Path $basePath))  
    {  
        $null = New-Item $basePath –Force  
    }  
 
    Set-ItemProperty $basePath -Name EnableScriptBlockInvocationLogging -Value "1" 
}

Function Enable-PSTranscription 
{  
    [CmdletBinding()]  
    param(  
        $OutputDirectory,  
        [Switch] $IncludeInvocationHeader  
    )  
 
    ## Ensure the base path exists  
    $basePath = "HKLM:\Software\Policies\Microsoft\Windows\PowerShell\Transcription" 
    if(-not (Test-Path $basePath))  
    {  
        $null = New-Item $basePath –Force  
    }

 

    ## Enable transcription  
    Set-ItemProperty $basePath -Name EnableTranscripting -Value 1
 

    ## Set the output directory  
    if($PSCmdlet.MyInvocation.BoundParameters.ContainsKey("OutputDirectory"))  
    {  
        Set-ItemProperty $basePath -Name OutputDirectory -Value $OutputDirectory  
    }
 

    ## Set the invocation header  
    if($IncludeInvocationHeader)  
    {  
        Set-ItemProperty $basePath -Name EnableInvocationHeader -Value 1  
    } 
}

<#Function Enable-ProtectedEventLogging 
{  
    param(  
        [Parameter(Mandatory)]  
        $Certificate  
    )
   
    $basePath = "HKLM:\Software\Policies\Microsoft\Windows\EventLog\ProtectedEventLogging"  
    if(-not (Test-Path $basePath))  
    {  
        $null = New-Item $basePath –Force  
    }
 

    Set-ItemProperty $basePath -Name EnableProtectedEventLogging -Value "1"  
    Set-ItemProperty $basePath -Name EncryptionCertificate -Value $Certificate
}
#>

#Enable-ProtectedEventLogging
Enable-PSTranscription
Enable-PSScriptBlockLogging
Function Set-RDPDisable
{
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
Function Set-RDPEnable
{
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
Function Set-RemoteDesktop
{
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
Function Set-RemoteRegistryKey {
    <#
    .SYNOPSIS
        Set registry key on remote computers.
    .DESCRIPTION
        This Function uses .Net class [Microsoft.Win32.RegistryKey].
    .PARAMETER ComputerName
        Name of the remote computers.
    .PARAMETER Hive
        Hive where the key is.
    .PARAMETER KeyPath
        Path of the key.
    .PARAMETER Name
        Name of the key setting.
    .PARAMETER Type
        Type of the key setting.
    .PARAMETER Value
        Value tu put in the key setting.
    .EXAMPLE
        Set-RemoteRegistryKey -ComputerName $env:ComputerName -Hive "LocalMachine" -KeyPath "software\corporate\master\Test" -Name "TestName" -Type String -Value "TestValue" -Verbose
    .LINK
        http://itfordummies.net
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$false,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$True,
            HelpMessage='Provide a ComputerName')]
        [String[]]$ComputerName=$env:ComputerName,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryHive]$Hive,
        
        [Parameter(Mandatory=$true)]
        [String]$KeyPath,
        
        [Parameter(Mandatory=$true)]
        [String]$Name,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryValueKind]$Type,
        
        [Parameter(Mandatory=$true)]
        [Object]$Value
    )
    Begin{
    }
    Process{
        ForEach ($Computer in $ComputerName) {
            try {
                Write-Verbose "Trying computer $Computer"
                $reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("$hive", "$Computer")
                Write-Debug -Message "Contenur de Reg $reg"
                $key=$reg.OpenSubKey("$KeyPath",$true)
                if($key -eq $null){
                    Write-Verbose -Message "Key not found."
                    Write-Verbose -Message "Calculating parent and child paths..."
                    $parent = Split-Path -Path $KeyPath -Parent
                    $child = Split-Path -Path $KeyPath -Leaf
                    Write-Verbose -Message "Creating the subkey $child in $parent..."
                    $Key=$reg.OpenSubKey("$parent",$true)
                    $Key.CreateSubKey("$child") | Out-Null
                    Write-Verbose -Message "Setting $value in $KeyPath"
                    $key=$reg.OpenSubKey("$KeyPath",$true)
                    $key.SetValue($Name,$Value,$Type)
                }
                else{
                    Write-Verbose "Key found, setting $Value in $KeyPath..."
                    $key.SetValue($Name,$Value,$Type)
                }
                Write-Verbose "$Computer done."
            }#End Try
            catch {Write-Warning "$Computer : $_"} 
        }#End ForEach
    }#End Process
    End{
    }
}
Function Set-ScreenResolution { 
 
<# 
    .Synopsis 
        Sets the Screen Resolution of the primary monitor 
    .Description 
        Uses Pinvoke and ChangeDisplaySettings Win32API to make the change 
    .Example 
        Set-ScreenResolution -Width 1024 -Height 768         
    #> 
param ( 
[Parameter(Mandatory=$true, 
           Position = 0)] 
[int] 
$Width, 
 
[Parameter(Mandatory=$true, 
           Position = 1)] 
[int] 
$Height 
) 
 
$pinvokeCode = @" 
 
using System; 
using System.Runtime.InteropServices; 
 
namespace Resolution 
{ 
 
    [StructLayout(LayoutKind.Sequential)] 
    public struct DEVMODE1 
    { 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmDeviceName; 
        public short dmSpecVersion; 
        public short dmDriverVersion; 
        public short dmSize; 
        public short dmDriverExtra; 
        public int dmFields; 
 
        public short dmOrientation; 
        public short dmPaperSize; 
        public short dmPaperLength; 
        public short dmPaperWidth; 
 
        public short dmScale; 
        public short dmCopies; 
        public short dmDefaultSource; 
        public short dmPrintQuality; 
        public short dmColor; 
        public short dmDuplex; 
        public short dmYResolution; 
        public short dmTTOption; 
        public short dmCollate; 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmFormName; 
        public short dmLogPixels; 
        public short dmBitsPerPel; 
        public int dmPelsWidth; 
        public int dmPelsHeight; 
 
        public int dmDisplayFlags; 
        public int dmDisplayFrequency; 
 
        public int dmICMMethod; 
        public int dmICMIntent; 
        public int dmMediaType; 
        public int dmDitherType; 
        public int dmReserved1; 
        public int dmReserved2; 
 
        public int dmPanningWidth; 
        public int dmPanningHeight; 
    }; 
 
 
 
    class User_32 
    { 
        [DllImport("user32.dll")] 
        public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE1 devMode); 
        [DllImport("user32.dll")] 
        public static extern int ChangeDisplaySettings(ref DEVMODE1 devMode, int flags); 
 
        public const int ENUM_CURRENT_SETTINGS = -1; 
        public const int CDS_UPDATEREGISTRY = 0x01; 
        public const int CDS_TEST = 0x02; 
        public const int DISP_CHANGE_SUCCESSFUL = 0; 
        public const int DISP_CHANGE_RESTART = 1; 
        public const int DISP_CHANGE_FAILED = -1; 
    } 
 
 
 
    public class PrmaryScreenResolution 
    { 
        static public string ChangeResolution(int width, int height) 
        { 
 
            DEVMODE1 dm = GetDevMode1(); 
 
            if (0 != User_32.EnumDisplaySettings(null, User_32.ENUM_CURRENT_SETTINGS, ref dm)) 
            { 
 
                dm.dmPelsWidth = width; 
                dm.dmPelsHeight = height; 
 
                int iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_TEST); 
 
                if (iRet == User_32.DISP_CHANGE_FAILED) 
                { 
                    return "Unable To Process Your Request. Sorry For This Inconvenience."; 
                } 
                else 
                { 
                    iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_UPDATEREGISTRY); 
                    switch (iRet) 
                    { 
                        case User_32.DISP_CHANGE_SUCCESSFUL: 
                            { 
                                return "Success"; 
                            } 
                        case User_32.DISP_CHANGE_RESTART: 
                            { 
                                return "You Need To Reboot For The Change To Happen.\n If You Feel Any Problem After Rebooting Your Machine\nThen Try To Change Resolution In Safe Mode."; 
                            } 
                        default: 
                            { 
                                return "Failed To Change The Resolution"; 
                            } 
                    } 
 
                } 
 
 
            } 
            else 
            { 
                return "Failed To Change The Resolution."; 
            } 
        } 
 
        private static DEVMODE1 GetDevMode1() 
        { 
            DEVMODE1 dm = new DEVMODE1(); 
            dm.dmDeviceName = new String(new char[32]); 
            dm.dmFormName = new String(new char[32]); 
            dm.dmSize = (short)Marshal.SizeOf(dm); 
            return dm; 
        } 
    } 
} 
 
"@ 
 
Add-Type $pinvokeCode -ErrorAction SilentlyContinue 
[Resolution.PrmaryScreenResolution]::ChangeResolution($width,$height) 
} 
#md c:\Transcripts
 

## Kill all inherited permissions
$acl = Get-Acl c:\Transcripts
$acl.SetAccessRuleProtection($true, $false)
 

## Grant Administrators full control
$administrators = [System.Security.Principal.NTAccount] "Administrators"
$permission = $administrators,"FullControl","ObjectInherit,ContainerInherit","None","Allow"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.AddAccessRule($accessRule)
 

## Grant everyone else Write and ReadAttributes. This prevents users from listing
## transcripts from other machines on the domain.
$everyone = [System.Security.Principal.NTAccount] "Everyone"
$permission = $everyone,"Write,ReadAttributes","ObjectInherit,ContainerInherit","None","Allow"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.AddAccessRule($accessRule)
 

## Deny "Creator Owner" everything. This prevents users from
## viewing the content of previously written files.
$creatorOwner = [System.Security.Principal.NTAccount] "Creator Owner"
$permission = $creatorOwner,"FullControl","ObjectInherit,ContainerInherit","InheritOnly","Deny"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.AddAccessRule($accessRule)
 

## Set the ACL
$acl | Set-Acl c:\Transcripts\
 

## Create the SMB Share, granting Everyone the right to read and write files. Specific
## actions will actually be enforced by the ACL on the file folder.
#New-SmbShare -Name Transcripts -Path c:\Transcripts -ChangeAccess Everyone 
Function Set-VMEvcMode {
<#  
.SYNOPSIS  
    Configures the EVC status of a VM
.DESCRIPTION 
    Will configure the EVC status for the specified VM
.NOTES  
    Author:  Kyle Ruddy, @kmruddy, thatcouldbeaproblem.com
.PARAMETER Name
    VM name which the Function should be ran against
.PARAMETER EvcMode
    The EVC Mode key which should be set
.EXAMPLE
	Set-VMEvcMode -Name vmName -EvcMode intel-sandybridge
	Configures the EVC status of the provided VM to be 'intel-sandybridge'
#>
[CmdletBinding()] 
	param(
	[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        $Name,
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateSet("intel-merom","intel-penryn","intel-nehalem","intel-westmere","intel-sandybridge","intel-ivybridge","intel-haswell","intel-broadwell","intel-skylake","amd-rev-e","amd-rev-f","amd-greyhound-no3dnow","amd-greyhound","amd-bulldozer","amd-piledriver","amd-steamroller","amd-zen")]
        $EvcMode
  	)

    Process {
        $evVM = @()
        $updateVM = @()

        if ($name -is [string]) {$evVM += Get-VM -Name $Name -ErrorAction SilentlyContinue}
        elseif ($name -is [array]) {

            if ($name[0] -is [string]) {
                $name | foreach {
                    $evVM += Get-VM -Name $_ -ErrorAction SilentlyContinue
                }
            }
            elseif ($name[0] -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM = $name}

        }
        elseif ($name -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM += $name}
        
        if ($evVM -eq $null) {Write-Warning "No VMs found."}
        else {

            $si = Get-View ServiceInstance
            $evcMask = $si.Capability.SupportedEvcMode | where-object {$_.key -eq $EvcMode} | select -ExpandProperty FeatureMask

            foreach ($v in $evVM) {

                if (($v.HardwareVersion -ge 'vmx-14' -and $v.PowerState -eq 'PoweredOff') -or ($v.Version -ge 'v14' -and $v.PowerState -eq 'PoweredOff')) {

                    $v.ExtensionData.ApplyEvcModeVM_Task($evcMask, $true) | Out-Null
                    $updateVM += $v.Name
                                        
                }
                else {Write-Warning $v.Name + " does not have the minimum requirements of being Hardware Version 14 and powered off."}

            }

            if ($updateVM) {
            
            Start-Sleep -Seconds 2
            Get-VMEvcMode -Name $updateVM
            
            }

        }

    }

}
Function _SetDocumentProperty 
{
	#jeff hicks
	Param(
		[object] $Properties,
		[string] $Name,
		[string] $Value
	)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}
Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' 
			{
				ForEach($Cell in $Collection) 
				{
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
				If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end Switch
	} # end process
}
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'zh-'	{ '???? 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}
Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this Functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#############################################################################
##
## Show-Object
##
## From Windows PowerShell Cookbook (O'Reilly)
## by Lee Holmes (http://www.leeholmes.com/guide)
##
##############################################################################
Function show-object
{
<#

.SYNOPSIS

Provides a graphical interface to let you explore and navigate an object.


.EXAMPLE

PS > $ps = { Get-Process -ID $pid }.Ast
PS > Show-Object $ps

#>

param(
    ## The object to examine
    [Parameter(ValueFromPipeline = $true)]
    $InputObject
)

Set-StrictMode -Version 3

Add-Type -Assembly System.Windows.Forms

## Figure out the variable name to use when displaying the
## object navigation syntax. To do this, we look through all
## of the variables for the one with the same object identifier.
$rootVariableName = dir variable:\* -Exclude InputObject,Args |
    Where-Object {
        $_.Value -and
        ($_.Value.GetType() -eq $InputObject.GetType()) -and
        ($_.Value.GetHashCode() -eq $InputObject.GetHashCode())
}

## If we got multiple, pick the first
$rootVariableName = $rootVariableName| % Name | Select -First 1

## If we didn't find one, use a default name
if(-not $rootVariableName)
{
    $rootVariableName = "InputObject"
}

## A Function to add an object to the display tree
Function PopulateNode($node, $object)
{
    ## If we've been asked to add a NULL object, just return
    if(-not $object) { return }

    ## If the object is a collection, then we need to add multiple
    ## children to the node
    if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($object))
    {
        ## Some very rare collections don't support indexing (i.e.: $foo[0]).
        ## In this situation, PowerShell returns the parent object back when you
        ## try to access the [0] property.
        $isOnlyEnumerable = $object.GetHashCode() -eq $object[0].GetHashCode()

        ## Go through all the items
        $count = 0
        foreach($childObjectValue in $object)
        {
            ## Create the new node to add, with the node text of the item and
            ## value, along with its type
            $newChildNode = New-Object Windows.Forms.TreeNode
            $newChildNode.Text = "$($node.Name)[$count] = $childObjectValue : " +
                $childObjectValue.GetType()

            ## Use the node name to keep track of the actual property name
            ## and syntax to access that property.
            ## If we can't use the index operator to access children, add
            ## a special tag that we'll handle specially when displaying
            ## the node names.
            if($isOnlyEnumerable)
            {
                $newChildNode.Name = "@"
            }

            $newChildNode.Name += "[$count]"
            $null = $node.Nodes.Add($newChildNode)               

            ## If this node has children or properties, add a placeholder
            ## node underneath so that the node shows a '+' sign to be
            ## expanded.
            AddPlaceholderIfRequired $newChildNode $childObjectValue

            $count++
        }
    }
    else
    {
        ## If the item was not a collection, then go through its
        ## properties
        foreach($child in $object.PSObject.Properties)
        {
            ## Figure out the value of the property, along with
            ## its type.
            $childObject = $child.Value
            $childObjectType = $null
            if($childObject)
            {
                $childObjectType = $childObject.GetType()
            }

            ## Create the new node to add, with the node text of the item and
            ## value, along with its type
            $childNode = New-Object Windows.Forms.TreeNode
            $childNode.Text = $child.Name + " = $childObject : $childObjectType"
            $childNode.Name = $child.Name
            $null = $node.Nodes.Add($childNode)

            ## If this node has children or properties, add a placeholder
            ## node underneath so that the node shows a '+' sign to be
            ## expanded.
            AddPlaceholderIfRequired $childNode $childObject
        }
    }
}

## A Function to add a placeholder if required to a node.
## If there are any properties or children for this object, make a temporary
## node with the text "..." so that the node shows a '+' sign to be
## expanded.
Function AddPlaceholderIfRequired($node, $object)
{
    if(-not $object) { return }

    if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($object) -or
        @($object.PSObject.Properties))
    {
        $null = $node.Nodes.Add( (New-Object Windows.Forms.TreeNode "...") )
    }
}

## A Function invoked when a node is selected.
Function OnAfterSelect
{
    param($Sender, $TreeViewEventArgs)

    ## Determine the selected node
    $nodeSelected = $Sender.SelectedNode

    ## Walk through its parents, creating the virtual
    ## PowerShell syntax to access this property.
    $nodePath = GetPathForNode $nodeSelected

    ## Now, invoke that PowerShell syntax to retrieve
    ## the value of the property.
    $resultObject = Invoke-Expression $nodePath
    $outputPane.Text = $nodePath

    ## If we got some output, put the object's member
    ## information in the text box.
    if($resultObject)
    {
        $members = Get-Member -InputObject $resultObject | Out-String       
        $outputPane.Text += "`n" + $members
    }
}

## A Function invoked when the user is about to expand a node
Function OnBeforeExpand
{
    param($Sender, $TreeViewCancelEventArgs)

    ## Determine the selected node
    $selectedNode = $TreeViewCancelEventArgs.Node

    ## If it has a child node that is the placeholder, clear
    ## the placeholder node.
    if($selectedNode.FirstNode -and
        ($selectedNode.FirstNode.Text -eq "..."))
    {
        $selectedNode.Nodes.Clear()
    }
    else
    {
        return
    }

    ## Walk through its parents, creating the virtual
    ## PowerShell syntax to access this property.
    $nodePath = GetPathForNode $selectedNode 

    ## Now, invoke that PowerShell syntax to retrieve
    ## the value of the property.
    Invoke-Expression "`$resultObject = $nodePath"

    ## And populate the node with the result object.
    PopulateNode $selectedNode $resultObject
}

## A Function to handle keypresses on the form.
## In this case, we capture ^C to copy the path of
## the object property that we're currently viewing.
Function OnKeyPress
{
    param($Sender, $KeyPressEventArgs)

    ## [Char] 3 = Control-C
    if($KeyPressEventArgs.KeyChar -eq 3)
    {
        $KeyPressEventArgs.Handled = $true

        ## Get the object path, and set it on the clipboard
        $node = $Sender.SelectedNode
        $nodePath = GetPathForNode $node
        [System.Windows.Forms.Clipboard]::SetText($nodePath)

        $form.Close()
    }
}

## A Function to walk through the parents of a node,
## creating virtual PowerShell syntax to access this property.
Function GetPathForNode
{
    param($Node)

    $nodeElements = @()

    ## Go through all the parents, adding them so that
    ## $nodeElements is in order.
    while($Node)
    {
        $nodeElements = ,$Node + $nodeElements
        $Node = $Node.Parent
    }

    ## Now go through the node elements
    $nodePath = ""
    foreach($Node in $nodeElements)
    {
        $nodeName = $Node.Name

        ## If it was a node that PowerShell is able to enumerate
        ## (but not index), wrap it in the array cast operator.
        if($nodeName.StartsWith('@'))
        {
            $nodeName = $nodeName.Substring(1)
            $nodePath = "@(" + $nodePath + ")"
        }
        elseif($nodeName.StartsWith('['))
        {
            ## If it's a child index, we don't need to
            ## add the dot for property access
        }
        elseif($nodePath)
        {
            ## Otherwise, we're accessing a property. Add a dot.
            $nodePath += "."
        }

        ## Append the node name to the path
        $nodePath += $nodeName
    }

    ## And return the result
    $nodePath
}

## Create the TreeView, which will hold our object navigation
## area.
$treeView = New-Object Windows.Forms.TreeView
$treeView.Dock = "Top"
$treeView.Height = 500
$treeView.PathSeparator = "."
$treeView.Add_AfterSelect( { OnAfterSelect @args } )
$treeView.Add_BeforeExpand( { OnBeforeExpand @args } )
$treeView.Add_KeyPress( { OnKeyPress @args } )

## Create the output pane, which will hold our object
## member information.
$outputPane = New-Object System.Windows.Forms.TextBox
$outputPane.Multiline = $true
$outputPane.ScrollBars = "Vertical"
$outputPane.Font = "Consolas"
$outputPane.Dock = "Top"
$outputPane.Height = 300

## Create the root node, which represents the object
## we are trying to show.
$root = New-Object Windows.Forms.TreeNode
$root.Text = "$InputObject : " + $InputObject.GetType()
$root.Name = '$' + $rootVariableName
$root.Expand()
$null = $treeView.Nodes.Add($root)

## And populate the initial information into the tree
## view.
PopulateNode $root $InputObject

## Finally, create the main form and show it.
$form = New-Object Windows.Forms.Form
$form.Text = "Browsing " + $root.Text
$form.Width = 1000
$form.Height = 800
$form.Controls.Add($outputPane)
$form.Controls.Add($treeView)
$null = $form.ShowDialog()
$form.Dispose()
}
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


Start-RDP -server "ServerIp"
Function Stop-WinWord
{
	Write-Debug "***Enter Stop-WinWord"
	
	## determine our login session
	$proc = Get-Process -PID $PID
	If( $null -eq $proc )
	{
		throw "Stop-WinWord: Cannot find process $PID"
	}
	
	$SessionID = $proc.SessionId
	If( $null -eq $SessionID )
	{
		Write-Debug "Stop-WinWord: SessionId on $PID is null"
		throw "Can't find a session for pid $PID"
	}

	If( 0 -eq $SessionID )
	{
		Write-Debug "Stop-WinWord: SessionId is 0 -- that is a bug"
		throw "SessionId is zero for pid $PID"
	}
	
	#Find out if winword is running in our session
	try 
	{
		$wordProc = Get-Process 'WinWord' -ErrorAction SilentlyContinue
	}
	catch
	{
		Write-Debug "***Exit Stop-WinWord: no WinWord tasks are running #1"
		Return ## not running
	}

	If( !$wordproc )
	{
		Write-Debug "***Exit Stop-WinWord: no WinWord tasks are running #2"
		Return ## WinWord is not running in ANY session
	}
	
	$wordrunning = $wordProc |? { $_.SessionId -eq $SessionID }
	If( !$wordrunning )
	{
		Write-Debug "***Exit Stop-WinWord: wordRunning eq null"
		Return ## not running in the current session
	}
	If( $wordrunning -is [Array] )
	{
		Write-Debug "***Exit Stop-WinWord: wordRunning is an array, elements=$($wordrunning.Count)"
		throw "Multiple Word processes are running in session $SessionID"
	}

	## it is possible for the below to throw a fault if Winword stops before it is executed.
	Stop-Process -Id $wordrunning.Id -ErrorAction SilentlyContinue
	Write-Debug "***Exit Stop-WinWord: sent Stop-Process to $($wordrunning.Id)"
}
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
Function Test-DatePattern
{
#http://jdhitsolutions.com/blog/2014/10/powershell-dates-times-and-formats/
$patterns = "d","D","g","G","f","F","m","o","r","s", "t","T","u","U","Y","dd","MM","yyyy","yy","hh","mm","ss","yyyyMMdd","yyyyMMddhhmm","yyyyMMddhhmmss"

Write-host "It is now $(Get-Date)" -ForegroundColor Green

foreach ($pattern in $patterns) {

#create an Object
[pscustomobject]@{
 Pattern = $pattern
 Syntax = "Get-Date -format '$pattern'"
 Value = (Get-Date -Format $pattern)
}

} #foreach

Write-Host "Most patterns are case sensitive" -ForegroundColor Green
}
## Domain Trust DNS Testing
$IPList = '192.168.1.4','192.168.1.10','10.201.2.10','172.20.0.5','172.20.0.6'
$FQDN = 'USONVSVRDC01.USON.LOCAL','USONVSVRDC02.USON.LOCAL','USONVSVRDC03.USON.LOCAL','CLFRDC01.cloud.local','CLFRDC02.cloud.local'
$CN = 'USONVSVRDC01','USONVSVRDC02','USONVSVRDC03','CLFRDC01','CLFRDC02'

ForEach($IP in $IPList){
    if (test-connection $IP -quiet) { write-output "$IP Alive" }
        else
      { write-output "$IP Not Responding" }}
ForEach($fq in $FQDN){
    if (test-connection $fq -quiet) { write-output "$fq Alive" }
        else
      { write-output "$fq Not Responding" }}
ForEach($name in $CN){
    if (test-connection $name -quiet) { write-output "$name Alive" }
        else
      { write-output "$name Not Responding" }}
Function Test-IsLocalAdministrator
{
<#
.SYNOPSIS
	Function to verify if the current user is a local Administrator on the current system
.DESCRIPTION
	Function to verify if the current user is a local Administrator on the current system
.EXAMPLE
	Test-IsLocalAdministrator

	True
.NOTES
	Francois-Xavier Cat
	@lazywinadm
	www.lazywinadmin.com
	github.com/lazywinadmin
#>
	([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
Function Test-RegistryKey {
<#
.SYNOPSIS
Tests if a registry Key and/or value exists. Can also test if the value is of certain type
.DESCRIPTION
Based on PSADT’s Set-RegistryKey Function
Since registry values are treated like properties to a registry key so they have no specific path and you cannot use Test-Path to check whether a given registry value exists.
This provides a reliable way to check registry values even when it is empty or null.
Checks for registry keys
Can also do simple Value content -eq compare to save you from writing another If statement
Uses Write-Log to aid troubleshooting
.PARAMETER Key
Path of the registry key (Required).
.PARAMETER Name
The value name (optional).
.PARAMETER Value
Value to compare against (optional).
This is to save you from writing another If statement if doing a simple -eq test
If comparing ‘ExpandString’, Pre-Expand the value first
If comparing ‘Binary’, see example
If comparing ‘MultiString’, please report how you did it if successful
.PARAMETER Type
The type of registry value to create or set. Options: ‘Binary’,’DWord’,’MultiString’,’QWord’,’String’. (optional).
CAVEAT: Unable to test for ExpandString. Current test will say a value of type ExpandString is a String
UN-supported Types: ‘ExpandString’,’None’,’Unknown’

.EXAMPLE
Key exists tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full”	#Key exists in win7	(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP”	#(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Do Not Exist”	#(returns $false)

Value exist tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full” -Name Install	#(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘DoesNotExist’ #(returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘IMblank’ #=<blank string>(returns $true)

Value Type tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full” -Name Install -Type DWord	#REG_SZ (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘IMblank’ -Type String	#=<blank string>(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘IMblank’ -Type MultiString	#=<blank string>(returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment” -Name Temp -Type DWord	#REG_EXP_SZ (returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment” -Name Temp -Type String	#REG_EXP_SZ (returns $true)

(Default) Value tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘(Default)’ #=Not set (returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4.0” -Name ‘(Default)’	#=deprecated	(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4.0” -Name ‘(Default)’ -Type ‘String’	#=deprecated (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4.0” -Name ‘(Default)’ -Type ‘MultiString’	#=deprecated (returns $false)

Value Content tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full” -Name Install -Value 1	#REG_DWORD=1 returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full” -Name Install -Value “1”	#REG_DWORD=1 returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5” -Name InstallPath -Value “C:\Windows\Microsoft.NET\Framework64\v3.5\” #REG_SZ (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5” -Name Version -Value “3.5.30729.5420”	#REG_SZ (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment” -Name ComSpec -Value “%SystemRoot%\system32\cmd.exe”	#REG_EXP_SZ (returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment” -Name ComSpec -Value “C:\Windows\system32\cmd.exe”	#REG_SZ (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management” -Name “ExistingPageFiles” -Value “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management” #REG_MULTI_SZ (returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power” -Name SystemPowerPolicy -Value “010101” #REZ_BINARY (returns $false)
$test=”1 0 0 0 2 0 0 0 0 0 0 0 0 0 0 0 2 0 0 0 0 0 0 0 0 0 0 0 2 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 2 0 0 0 0 0 0 0 0 0 0 0 16 14 0 0 90 0 0 0 4 0 0 0 4 0 0 0 1 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 10 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 176 4 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 176 4 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0?
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power” -Name SystemPowerPolicy -Value $test #returns $true, but machine specific
.NOTES
Limits:
a (default) value that is “Not Set” = “does not exist” and will return $False
Cannot test for both ValueData and Type in the same Function call
#>
[CmdletBinding()]
param(
[Parameter(Mandatory=$true,HelpMessage = “Path of the registry key”)]
[ValidateNotNullorEmpty()]
[string]$Key,
[Parameter(Mandatory=$false,HelpMessage = “The value name (optional). For a key’s default value use (default)”)]
[ValidateNotNull()]
[String]$Name,
[Parameter(Mandatory=$false,HelpMessage = “Value Type (optional). Use ‘String’ for ‘ExpandString'”)]
[ValidateNotNullOrEmpty()]
[ValidateSet(‘Binary’,’DWord’,’MultiString’,’QWord’,’String’)]
[Microsoft.Win32.RegistryValueKind]$Type,
[Parameter(Mandatory=$false)]
$Value	#do NOT cast data type! $null or empty could also be given!
)

Begin {
## Get the name of this Function and write header
[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
}
Process {
$Key = Convert-RegistryPath -Key $Key

if( -not (Test-Path -Path $Key -PathType Container) ) {
Write-Host “Key [$Key] does NOT Exists” -Source ${CmdletName}
return $false #no point testing for the value
}

$AllKeyValues = Get-ItemProperty -Path $Key
If ($PSBoundParameters.ContainsKey(‘ValueName’) ) {	#	If ($Name) {
if ( -not $AllKeyValues ) {
Write-log “Key [$Key] Exists but has no values!” -Source ${CmdletName}
return $false
}

$ValDataRead = Get-Member -InputObject $AllKeyValues -Name $Name	#CAVEAT: converts REG_SZ to Int32 if it’s a number
Write-log “Value [$Name] Exists and contains [$ValDataRead]” -Source ${CmdletName}
if( $ValDataRead ) {
$ValDataRead = $($AllKeyValues.$Name)	#CAVEAT: converts REG_SZ to Int32 if it’s a number
Write-log “Value [$Name] Exists and contains [$ValDataRead]” -Source ${CmdletName}

If ($Value) { # do a data compare (Why not, we have everything else)
#	$ValDataRead = Get-ItemProperty -Path $Key -Name $Name	#CAVEAT: returns hash array @{Install=1}
#	Write-log “Value [$Name] Exists and contains [$ValDataRead]”
If ($Value -eq $ValDataRead) {	#If there is no way to
Return $true
} Else { Return $false }
} ElseIf ($Type) {	# do a Type compare (Why not, we have everything else)
$ValTypeRead = switch ($Name.gettype().Name) {
“String”{‘String’}	# {“REG_SZ”; } # also REG_EXP_SZ [ hex(2) ]
“Int32” {‘DWord’}	# {“REG_DWORD”; }
“Int64” {‘QWord’}	# {“REG_QWORD”; }
“String[]” {‘MultiString’}# {“REG_MULTI_SZ”; }
“Byte[]” {‘Binary’}	# {“REG_BINARY”}
default {‘Unknown’}
}
Write-log “Value [$Name] is of type [$ValTypeRead]” -Source ${CmdletName}
If ($ValTypeRead -eq $Type) {
Return $true
} else {return $false }

} Else {
return $true
} # If ($Value) {
} else {
Write-log “Key [$Key] exist but [$Name] does not exist” -Source ${CmdletName}
return $false
}
} Else { #The Key exists but we don’t care about its values
Write-log “Key [$Key] Exists” -Source ${CmdletName}
return $true
}
}
End {
Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
}

}	# Test-RegistryKey
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}
Function Test-RemoteDesktopIsEnabled
{
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
  
}#Function
Function Test-ServerSSLSupport {
[CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$HostName,
        [UInt16]$Port = 443
    )
    process {
        $RetValue = New-Object psobject -Property @{
            Host = $HostName
            Port = $Port
            SSLv2 = $false
            SSLv3 = $false
            TLSv1_0 = $false
            TLSv1_1 = $false
            TLSv1_2 = $false
            KeyExhange = $null
            HashAlgorithm = $null
        }
        "ssl2", "ssl3", "tls", "tls11", "tls12" | %{
            $TcpClient = New-Object Net.Sockets.TcpClient
            $TcpClient.Connect($RetValue.Host, $RetValue.Port)
            $SslStream = New-Object Net.Security.SslStream $TcpClient.GetStream()
            $SslStream.ReadTimeout = 15000
            $SslStream.WriteTimeout = 15000
            try {
                $SslStream.AuthenticateAsClient($RetValue.Host,$null,$_,$false)
                $RetValue.KeyExhange = $SslStream.KeyExchangeAlgorithm
                $RetValue.HashAlgorithm = $SslStream.HashAlgorithm
                $status = $true
            } catch {
                $status = $false
            }
            switch ($_) {
                "ssl2" {$RetValue.SSLv2 = $status}
                "ssl3" {$RetValue.SSLv3 = $status}
                "tls" {$RetValue.TLSv1_0 = $status}
                "tls11" {$RetValue.TLSv1_1 = $status}
                "tls12" {$RetValue.TLSv1_2 = $status}
            }
            # dispose objects to prevent memory leaks
            $TcpClient.Dispose()
            $SslStream.Dispose()
        }
        $RetValue
    }
}
<#
 .DESCRIPTION
   Outputs the SSL protocols that the client is able to successfully use to connect to a server.
 
 .NOTES
 
   Copyright 2014 Chris Duck
   http://blog.whatsupduck.net
 
   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at
 
     http://www.apache.org/licenses/LICENSE-2.0
 
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
 
 .PARAMETER ComputerName
   The name of the remote computer to connect to.
 
 .PARAMETER Port
   The remote port to connect to. The default is 443.
 
 .EXAMPLE
   Test-SslProtocols -ComputerName "www.google.com"
   
   ComputerName       : www.google.com
   Port               : 443
   KeyLength          : 2048
   SignatureAlgorithm : rsa-sha1
   Ssl2               : False
   Ssl3               : True
   Tls                : True
   Tls11              : True
   Tls12              : True
 #>
 Function Test-SslProtocols {
   param(
     [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
     $ComputerName,
     
     [Parameter(ValueFromPipelineByPropertyName=$true)]
     [int]$Port = 443
   )
   begin {
     $ProtocolNames = [System.Security.Authentication.SslProtocols] | gm -static -MemberType Property | ?{$_.Name -notin @("Default","None")} | %{$_.Name}
   }
   process {
     $ProtocolStatus = [Ordered]@{}
     $ProtocolStatus.Add("ComputerName", $ComputerName)
     $ProtocolStatus.Add("Port", $Port)
     $ProtocolStatus.Add("KeyLength", $null)
     $ProtocolStatus.Add("SignatureAlgorithm", $null)
     
     $ProtocolNames | %{
       $ProtocolName = $_
       $Socket = New-Object System.Net.Sockets.Socket([System.Net.Sockets.SocketType]::Stream, [System.Net.Sockets.ProtocolType]::Tcp)
       $Socket.Connect($ComputerName, $Port)
       try {
         $NetStream = New-Object System.Net.Sockets.NetworkStream($Socket, $true)
         $SslStream = New-Object System.Net.Security.SslStream($NetStream, $true)
         $SslStream.AuthenticateAsClient($ComputerName,  $null, $ProtocolName, $false )
         $RemoteCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]$SslStream.RemoteCertificate
         $ProtocolStatus["KeyLength"] = $RemoteCertificate.PublicKey.Key.KeySize
         $ProtocolStatus["SignatureAlgorithm"] = $RemoteCertificate.PublicKey.Key.SignatureAlgorithm.Split("#")[1]
         $ProtocolStatus.Add($ProtocolName, $true)
       } catch  {
         $ProtocolStatus.Add($ProtocolName, $false)
       } finally {
         $SslStream.Close()
       }
     }
     [PSCustomObject]$ProtocolStatus
   }
 }
Function Send-Email
{
<#
	.SYNOPSIS
		This Function allows you to send email
	
	.DESCRIPTION
		This Function allows you to send email using the NET Class System.Net.Mail
	
	.PARAMETER To
		A description of the To parameter.
	
	.PARAMETER From
		A description of the From parameter.
	
	.PARAMETER FromDisplayName
		Specifies the DisplayName to show for the FROM parameter
	
	.PARAMETER SenderAddress
		A description of the SenderAddress parameter.
	
	.PARAMETER SenderDisplayName
		Specifies the DisplayName of the Sender
	
	.PARAMETER CC
		A description of the CC parameter.
	
	.PARAMETER BCC
		A description of the BCC parameter.
	
	.PARAMETER ReplyToList
		Specifies the email address(es) that will be use when the recipient(s) reply to the email.
	
	.PARAMETER Subject
		Specifies the subject of the email.
	
	.PARAMETER Body
		Specifies the body of the email.
	
	.PARAMETER BodyIsHTML
		Specifies that the text format of the body is HTML. Default is Plain Text.
	
	.PARAMETER Priority
		Specifies the priority of the message. Default is Normal.
	
	.PARAMETER Encoding
		Specifies the text encoding of the title and the body.
	
	.PARAMETER Attachment
		Specifies if an attachement must be added to the Function
	
	.PARAMETER Credential
		Specifies the credential to use, default will use the current credential.
	
	.PARAMETER SMTPServer
		Specifies if the SMTP Server IP or FQDN to use
	
	.PARAMETER Port
		Specifies if the SMTP Server Port to use. Default is 25.
	
	.PARAMETER EnableSSL
		Specifies if the email must be sent using SSL.
	
	.PARAMETER DeliveryNotificationOptions
		Specifies the delivey notification options.
		https://msdn.microsoft.com/en-us/library/system.net.mail.deliverynotificationoptions.aspx
	
	.PARAMETER EmailCC
		Specifies the Carbon Copy recipient
	
	.PARAMETER EmailBCC
		Specifies the Blind Carbon Copy recipient
	
	.PARAMETER EmailTo
		Specifies the recipient of the email
	
	.PARAMETER EmailFrom
		Specifies the sender of the email
	
	.PARAMETER Sender
		Specifies the Sender Email address. Sender is the Address of the actual sender acting on behalf of the author listed in the From parameter.
	
	.EXAMPLE
		Send-email `
		-EmailTo "fxcat@contoso.com" `
		-EmailFrom "powershell@contoso.com" `
		-SMTPServer "smtp.sendgrid.net"  `
		-Subject "Test Email" `
		-Body "Test Email"
		
		This will send an email using the current credential of the current logged user
	
	.EXAMPLE
		$Cred = [System.Net.NetworkCredential](Get-Credential -Credential testuser)
		
		Send-email `
		-EmailTo "fxcat@contoso.com" `
		-EmailFrom "powershell@contoso.com" `
		-Credential $cred
		-SMTPServer "smtp.sendgrid.net"  `
		-Subject "Test Email" `
		-Body "Test Email"
		
		This will send an email using the credentials specified in the $Cred variable
	
	.EXAMPLE
		Send-email `
		-EmailTo "fxcat@contoso.com","SomeoneElse@contoso.com" `
		-EmailFrom "powershell@contoso.com" `
		-SMTPServer "smtp.sendgrid.net"  `
		-Subject "Test Email" `
		-Body "Test Email"
		
		This will send an email using the current credential of the current logged user to two
		fxcat@contoso.com and SomeoneElse@contoso.com
	
	.NOTES
		Francois-Xavier Cat
		fxcat@lazywinadmin.com
		www.lazywinadmin.com
		@lazywinadm
		
		VERSION HISTORY
		1.0 2014/12/25 	Initial Version
		1.1 2015/02/04 	Adding some error handling and clean up the code a bit
		Add Encoding, CC, BCC, BodyAsHTML
		1.2 2015/04/02	Credential
		
		TODO
		-Add more Help/Example
		-Add Support for classic Get-Credential
#>
	
	[CmdletBinding(DefaultParameterSetName = 'Main')]
	param
	(
		[Parameter(ParameterSetName = 'Main',
				   Mandatory = $true)]
		[Alias('EmailTo')]
		[String[]]$To,
		
		[Parameter(ParameterSetName = 'Main',
				   Mandatory = $true)]
		[Alias('EmailFrom', 'FromAddress')]
		[String]$From,
		
		[Parameter(ParameterSetName = 'Main')]
		[ValidateNotNullOrEmpty()]
		[string]$FromDisplayName,
		
		[Parameter(ParameterSetName = 'Main')]
		[Alias('EmailCC')]
		[String]$CC,
		
		[Parameter(ParameterSetName = 'Main')]
		[Alias('EmailBCC')]
		[System.String]$BCC,
		
		[Parameter(ParameterSetName = 'Main')]
		[ValidateNotNullOrEmpty()]
		[Alias('ReplyTo')]
		[System.string[]]$ReplyToList,
		
		[Parameter(ParameterSetName = 'Main')]
		[System.String]$Subject = "Email from PowerShell",
		
		[Parameter(ParameterSetName = 'Main')]
		[System.String]$Body = "Hello World",
		
		[Parameter(ParameterSetName = 'Main')]
		[Switch]$BodyIsHTML = $false,
		
		[Parameter(ParameterSetName = 'Main')]
		[ValidateNotNullOrEmpty()]
		[System.Net.Mail.MailPriority]$Priority = "Normal",
		
		[Parameter(ParameterSetName = 'Main')]
		[ValidateSet("Default", "ASCII", "Unicode", "UTF7", "UTF8", "UTF32")]
		[System.String]$Encoding = "Default",
		
		[Parameter(ParameterSetName = 'Main')]
		[System.String]$Attachment,
		
		[Parameter(ParameterSetName = 'Main')]
		[System.Net.NetworkCredential]$Credential,
		
		[Parameter(ParameterSetName = 'Main',
				   Mandatory = $true)]
		[ValidateScript({
			# Verify the host is reachable
			Test-Connection -ComputerName $_ -Count 1 -Quiet
		})]
		[Alias("Server")]
		[string]$SMTPServer,
		
		[Parameter(ParameterSetName = 'Main')]
		[ValidateRange(1, 65535)]
		[Alias("SMTPServerPort")]
		[int]$Port = 25,
		
		[Parameter(ParameterSetName = 'Main')]
		[Switch]$EnableSSL,
		
		[Parameter(ParameterSetName = 'Main')]
		[ValidateNotNullOrEmpty()]
		[Alias('EmailSender', 'Sender')]
		[string]$SenderAddress,
		
		[Parameter(ParameterSetName = 'Main')]
		[ValidateNotNullOrEmpty()]
		[System.String]$SenderDisplayName,
		
		[Parameter(ParameterSetName = 'Main')]
		[ValidateNotNullOrEmpty()]
		[Alias('DeliveryOptions')]
		[System.Net.Mail.DeliveryNotificationOptions]$DeliveryNotificationOptions
	)
	
	#PARAM
	
	PROCESS
	{
		TRY
		{
			# Create Mail Message Object
			$SMTPMessage = New-Object -TypeName System.Net.Mail.MailMessage
			$SMTPMessage.From = $From
			FOREACH ($ToAddress in $To) { $SMTPMessage.To.add($ToAddress) }
			$SMTPMessage.Body = $Body
			$SMTPMessage.IsBodyHtml = $BodyIsHTML
			$SMTPMessage.Subject = $Subject
			$SMTPMessage.BodyEncoding = $([System.Text.Encoding]::$Encoding)
			$SMTPMessage.SubjectEncoding = $([System.Text.Encoding]::$Encoding)
			$SMTPMessage.Priority = $Priority
			$SMTPMessage.Sender = $SenderAddress
			
			# Sender Displayname parameter
			IF ($PSBoundParameters['SenderDisplayName'])
			{
				$SMTPMessage.Sender.DisplayName = $SenderDisplayName
			}
			
			# From Displayname parameter
			IF ($PSBoundParameters['FromDisplayName'])
			{
				$SMTPMessage.From.DisplayName = $FromDisplayName
			}
			
			# CC Parameter
			IF ($PSBoundParameters['CC'])
			{
				$SMTPMessage.CC.Add($CC)
			}
			
			# BCC Parameter
			IF ($PSBoundParameters['BCC'])
			{
				$SMTPMessage.BCC.Add($BCC)
			}
			
			# ReplyToList Parameter
			IF ($PSBoundParameters['ReplyToList'])
			{
				foreach ($ReplyTo in $ReplyToList)
				{
					$SMTPMessage.ReplyToList.Add($ReplyTo)
				}
			}
			
			# Attachement Parameter
			IF ($PSBoundParameters['attachment'])
			{
				$SMTPattachment = New-Object -TypeName System.Net.Mail.Attachment($attachment)
				$SMTPMessage.Attachments.Add($STMPattachment)
			}
			
			# Delivery Options
			IF ($PSBoundParameters['DeliveryNotificationOptions'])
			{
				$SMTPMessage.DeliveryNotificationOptions = $DeliveryNotificationOptions
			}
			
			#Create SMTP Client Object
			$SMTPClient = New-Object -TypeName Net.Mail.SmtpClient
			$SMTPClient.Host = $SmtpServer
			$SMTPClient.Port = $Port
			
			# SSL Parameter
			IF ($PSBoundParameters['EnableSSL'])
			{
				$SMTPClient.EnableSsl = $true
			}
			
			# Credential Paramenter
			#IF (($PSBoundParameters['Username']) -and ($PSBoundParameters['Password']))
			IF ($PSBoundParameters['Credential'])
			{
				<#
				# Create Credential Object
				$Credentials = New-Object -TypeName System.Net.NetworkCredential
				$Credentials.UserName = $username.Split("@")[0]
				$Credentials.Password = $Password
				#>
				
				# Add the credentials object to the SMTPClient obj
				$SMTPClient.Credentials = $Credential
			}
			IF (-not $PSBoundParameters['Credential'])
			{
				# Use the current logged user credential
				$SMTPClient.UseDefaultCredentials = $true
			}
			
			# Send the Email
			$SMTPClient.Send($SMTPMessage)
			
		}#TRY
		CATCH
		{
			Write-Warning -message "[PROCESS] Something wrong happened"
			Write-Warning -Message $Error[0].Exception.Message
		}
	}#Process
	END
	{
		# Remove Variables
		Remove-Variable -Name SMTPClient -ErrorAction SilentlyContinue
		Remove-Variable -Name Password -ErrorAction SilentlyContinue
	}#END
} #End Function Send-Email
Function Write-Log
{
<#
.SYNOPSIS
    Function to create or append a log file

#>
[CmdletBinding()]
    Param (
        [Parameter()]
        $Path="",
        $LogName = "$(Get-Date -f 'yyyyMMdd').log",
        
        [Parameter(Mandatory=$true)]
        $Message = "",

        [Parameter()]
        [ValidateSet('INFORMATIVE','WARNING','ERROR')]
        $Type = "INFORMATIVE",
        $Category
    )
    BEGIN {
        # Verify if the log already exists, else create it
        IF (-not(Test-Path -Path $(Join-Path -Path $Path -ChildPath $LogName))){
            New-Item -Path $(Join-Path -Path $Path -ChildPath $LogName) -ItemType file
        }
    
    }
    PROCESS{
        TRY{
            "$(Get-Date -Format yyyyMMdd:HHmmss) [$TYPE] [$category] $Message" | Out-File -FilePath (Join-Path -Path $Path -ChildPath $LogName) -Append
        }
        CATCH{
            Write-Error -Message "Could not write into $(Join-Path -Path $Path -ChildPath $LogName)"
            Write-Error -Message "Last Error:$($error[0].exception.message)"
        }
    }

}
# ---------------------------------------------------
# Version: 1.0
# Author: Joshua Duffney
# Date: 07/13/2014
# Description: Using PowerShell to Unblock files that are downloaded from the internet.
# Comments: Call the Function Unblcok with the path to the folder containing blocked files in "" after it, see line 15.
# ---------------------------------------------------

Function Unblock ($path) { 

Get-ChildItem "$path" -Recurse | Unblock-File

}
Get-Module -ListAvailable |
Where-Object ModuleBase -like $env:ProgramFiles\WindowsPowerShell\Modules\* |
Sort-Object -Property Name, Version -Descending |
Get-Unique -PipelineVariable Module |
ForEach-Object {
    if (-not(Test-Path -Path "$($_.ModuleBase)\PSGetModuleInfo.xml")) {
        Find-Module -Name $_.Name -OutVariable Repo -ErrorAction SilentlyContinue |
        Compare-Object -ReferenceObject $_ -Property Name, Version |
        Where-Object SideIndicator -eq '=>' |
        Select-Object -Property Name,
                                Version,
                                @{label='Repository';expression={$Repo.Repository}},
                                @{label='InstalledVersion';expression={$Module.Version}}
    }
} | ForEach-Object {Install-Module -Name $_.Name -Force}
<#
.SYNOPSIS
    A small wrapper for PowerShellGet to upgrade all installed modules and scripts.
.DESCRIPTION
    A small wrapper for PowerShellGet to upgrade all installed modules and scripts.
.PARAMETER WhatIf
    Show modules which would get upgraded.
.EXAMPLE
    Upgrade-InstalledModules.ps1

    Description
    -------------
    Updates modules installed via PowerShellGet.
.NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 5.0

       Version History
       1.0.0 - Initial release
#>
[CmdletBinding()]
Param (
    [Parameter(HelpMessage = 'Show modules which would get upgraded.')]
    [switch]$WhatIf
)

try {
    Import-Module PowerShellGet
}
catch {
    Write-Warning 'Unable to load PowerShellGet. This script only works with PowerShell 5 and greater.'
    return
}

$WhatIfParam = @{WhatIf=$WhatIf}

Get-InstalledModule | Update-Module -Force @WhatIfParam
#Get-InstalledScript | Update-Script @WhatIfParam
Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}
Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('???', '???', '??', '??', '??',
					'??(??)', '??(??)', '???', '??', '??(??)',
					'??(??)', '??', '??', '??', '???',
					'???')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}
Function View-Cats
{
	<#
	.SYNOPSIS
		This will open Internet explorer and show a different cat every 5 seconds
	.DESCRIPTION
	.NOTES
		#http://www.reddit.com/r/PowerShell/comments/2htfog/viewcats/
	#>
    Param(
        [int]$refreshtime=5
    )
    $IE = new-object -ComObject internetexplorer.application
    $IE.visible = $true
    $IE.FullScreen = $true
    $shell = New-Object -ComObject wscript.shell
    $shell.AppActivate("Internet Explorer")

    while($true){
        $request = Invoke-WebRequest -Uri "http://thecatapi.com/api/images/get" -Method get 
        $IE.Navigate($request.BaseResponse.ResponseUri.AbsoluteUri)
        Start-Sleep -Seconds $refreshtime
    }
} 
# Enable or Disable Hot Add Memory/CPU
# Enable-MemHotAdd $ServerName
# Disable-MemHotAdd $ServerName
# Enable-vCPUHotAdd $ServerName
# Disable-vCPUHotAdd $ServerName




Function Enable-MemHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-MemHotAdd($vm){
$vmview = Get-VM $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Enable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Get-VMEVCMode {
    <#  .Description
        Code to get VMs' EVC mode and that of the cluster in which the VMs reside.  May 2014, vNugglets.com
        .Example
        Get-VMEVCMode -Cluster myCluster | ?{$_.VMEVCMode -ne $_.ClusterEVCMode}
        Get all VMs in given clusters and return, for each, an object with the VM's- and its cluster's EVC mode, if any
        .Outputs
        PSCustomObject
    #>
    param(
        ## Cluster name pattern (regex) to use for getting the clusters whose VMs to get
        [string]$Cluster_str = ".+"
    )
 
    process {
        ## get the matching cluster View objects
        Get-View -ViewType ClusterComputeResource -Property Name,Summary -Filter @{"Name" = $Cluster_str} | Foreach-Object {
            $viewThisCluster = $_
            ## get the VMs Views in this cluster
            Get-View -ViewType VirtualMachine -Property Name,Runtime.PowerState,Summary.Runtime.MinRequiredEVCModeKey -SearchRoot $viewThisCluster.MoRef | Foreach-Object {
                ## create new PSObject with some nice info
                New-Object -Type PSObject -Property ([ordered]@{
                    Name = $_.Name
                    PowerState = $_.Runtime.PowerState
                    VMEVCMode = $_.Summary.Runtime.MinRequiredEVCModeKey
                    ClusterEVCMode = $viewThisCluster.Summary.CurrentEVCModeKey
                    ClusterName = $viewThisCluster.Name
                })
            } ## end foreach-object
        } ## end foreach-object
    } ## end process
} ## end Function
Get-ChildItem -path C:\ -recurse |dir | ? {$_.lastwritetime -gt '11/7/15' -AND $_.lastwritetime -lt '11/9/15'}|
    ft -AutoSize |
        out-file -filepath "c:\temp\modified.txt"-append
Function WMIDateStringToDate($Bootup) {    
    [System.Management.ManagementDateTimeconverter]::ToDateTime($Bootup)    
} 
 
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

#### Symantec Enterprise Protection ####
Function Get-SEPVersion {
<#
.SYNOPSIS
Retrieve Symantec Endpoint Version, Definition Date and Sylink Group

.DESCRIPTION
Retrieve Symantec Endpoint Version, Definition Date and Sylink Group

.PARAMETER  ComputerName
Name of the computer to query SEP info for

.EXAMPLE
PS C:\> Get-SEPVersion -ComputerName Server01

.EXAMPLE
PS C:\> $servers | Get-SEPVersion

.NOTES
Author: Jonathan Medd
Date: 23/12/2011
#>

[CmdletBinding()]
param(
[Parameter(Position=0,Mandatory=$true,HelpMessage="Name of the computer to query SEP for",
ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[Alias('CN','__SERVER','IPAddress','Server')]
[System.String]
$ComputerName
)

begin {
# Create object to enable access to the months of the year
$DateTimeFormat = New-Object System.Globalization.DateTimeFormatInfo

# Set Registry keys to query
$SMCKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
}


process {


try {

# Connect to Registry
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName)

# Obtain Product Version value
$SMCRegKey = $reg.opensubkey($SMCKey)
$SEPVersion = $SMCRegKey.GetValue('ProductVersion')

# Obtain Pattern File Date Value
$AVRegKey = $reg.opensubkey($AVKey)
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate')

# Convert PatternFileDate to readable date
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970)
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]]
$AVDayFileDate = [string]$AVPatternFileDate[2]
$AVFileVersionDate = $AVDayFileDate + " " + $AVMonthFileDate + " " + $AVYearFileDate

# Obtain Sylink Group value
$SylinkRegKey = $reg.opensubkey($SylinkKey)
$SylinkGroup = $SylinkRegKey.GetValue('CurrentGroup')

}

catch [System.Management.Automation.MethodInvocationException]

{
$SEPVersion = "Unable to connect to computer"
$AVFileVersionDate = ""
$SylinkGroup = ""
}

$MYObject = “” | Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate,SylinkGroup
$MYObject.ComputerName = $ComputerName
$MYObject.SEPProductVersion = $SEPVersion
$MYObject.SEPDefinitionDate = $AVFileVersionDate
$MYObject.SylinkGroup = $SylinkGroup
$MYObject

}
}
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
#update 5-May-2016 by Michael B. Smith
{
	Param(
		[int] $style       = 0, 
		[int] $tabs        = 0, 
		[string] $name     = '', 
		[string] $value    = '', 
		[string] $fontName = $null,
		[int] $fontSize    = 0,
		[bool] $italics    = $false,
		[bool] $boldface   = $false,
		[Switch] $nonewline
	)
	
	#Build output style
	[string]$output = ''
	Switch ($style)
	{
		0 {$Script:Selection.Style = $myHash.Word_NoSpacing}
		1 {$Script:Selection.Style = $myHash.Word_Heading1}
		2 {$Script:Selection.Style = $myHash.Word_Heading2}
		3 {$Script:Selection.Style = $myHash.Word_Heading3}
		4 {$Script:Selection.Style = $myHash.Word_Heading4}
		Default {$Script:Selection.Style = $myHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t" 
		$tabs-- 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If( !$nonewline )
	{
		$Script:Selection.TypeParagraph()
	}
}
