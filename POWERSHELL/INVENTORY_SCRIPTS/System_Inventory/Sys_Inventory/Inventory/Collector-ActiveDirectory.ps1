<#
    .SYNOPSIS
        Produce a XML collection file from "Active Directory" source. (v5.0.101 - May 02, 2017)

    .DESCRIPTION
        Creates artifact collection file for uploading through the Microsoft WorkSpace service for SAM and APEX
        engagements that require XML documents in a specific format. 

        Collector-ActiveDirectory gets required information (Computers or Users) from Active Directory and generates
        the XML artifact in the required name and format.

        The script is to be executed on a machine with access to the Active Directory repository.

        Note: To enable PowerShell script execution, it may be necessary to change the script execution policy.
              To change the execution policy, from the PowerShell command prompt type the following command and press Enter
              eg. PS C:>Set-ExecutionPolicy Unrestricted -Scope CurrentUser

    .PARAMETER  XmlFilePath
        Fully Qualified Path to storage location of XML file to be created.  The application will automatically generate the file name.

    .PARAMETER  FilterType
        This parameter specifies the type of filter to be applied.  The two valid values are "Computer" and "User".
        If no value is specified the file collection will default to "Computer".

    .PARAMETER  DomainName
        This parameter specifies the name of the domain to access.
        The machine running this script must be trusted on the specified domain and the current user account
        or supplied credentials must have permission to access the information on the specified domain.

    .PARAMETER  NumberOfDomainControllersToQuery
        Specifies how many Domain Controllers to query in the current Domain.
        If no value is specified for this parameter the script will default to querying only one Domain Controller.
        To query all Domain Controllers in the Domain, set this parameter to 0.

    .PARAMETER  SpecificDomainControllerToQuery
        If a value is supplied in this parameter, all Domain Controllers that begin
        with a matching value will be selected.  The more specific the value the finer
        the result set.  Also, if a value is supplied in this parameter any commandline
        value supplied for the NumberOfDomainControllersToQuery will be ignored.

    .PARAMETER  IntegratedSecurity
        If specified, the script will execute using the current users Domain credentials.
        If this value is not specified the user must either pass in a valid PSCredential object in the Credential parameter
        or if neither is specified the system will prompt the user to enter credentials at time execution.

    .PARAMETER  Credential
        This parameter accepts a PSCredential object. If this value is specified and cannot be validated the script will 
        fail even if the IntegratedSecurity parameter is also selected.  It is best to use one or the other when executing
        this script.

    .PARAMETER  ProgressDisplay
        If specified, the command window includes a progress activity indicator.

    .PARAMETER  SuppressLogFile
        Log files are created by default, if this switch is included the creation of a Log file will be suppressed.

    .PARAMETER  LogFilePath
        Fully Qualified Path to storage location of Log file to be created.  The application will automatically generate the file name.
        If no path is specified and the SuppressLogFile switch is not included, the Log file will created in the same folder as the Xml File.

    .PARAMETER  xDTCall
        This value is used for internal processing and should be ignored when running this script in a PowerShell command window.

    .PARAMETER  DataSource
        This value is used for internal processing and should be ignored when running this script in a PowerShell command window.

    .PARAMETER  AppVersion
        This value is used for internal processing and should be ignored when running this script in a PowerShell command window.

    .EXAMPLE
        C:\TEMP\Collector-ActiveDirectory.ps1 "C:\TEMP\"

    .EXAMPLE
        C:\TEMP\Collector-ActiveDirectory.ps1 -XmlFilePath "C:\TEMP\" -FilterType "User" -DomainName "MyDomain.local" -NumberOfDomainControllersToQuery 1 -SpecificDomainControllerToQuery "ControllerName" -IntegratedSecurity -SuppressLogFile -ProgressDisplay

    .LINK
        Author: Inviso Corporation
        Website: InvisoCorp.com/SAM
        Support Email: InvisoSA@InvisoCorp.com

    .NOTES
        DISCLAIMER: The sample scripts are not supported under any Microsoft standard support program or 
        service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further 
        disclaims all implied warranties including, without limitation, any implied warranties of merchantability
        or of fitness for a particular purpose. The entire risk arising out of the use or performance of 
        the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, 
        or anyone else involved in the creation, production, or delivery of the scripts be liable for any 
        damages whatsoever (including, without limitation, damages for loss of business profits, business 
        interruption, loss of business information, or other pecuniary loss) arising out of the use of or 
        inability to use the sample scripts or documentation, even if Microsoft has been advised of the 
        possibility of such damages.

#>
[CmdletBinding(SupportsShouldProcess=$true)]
Param (
    [Parameter(Mandatory=$true,Position=0)]
        [string] $XmlFilePath,
    [Parameter(Position=1)]
        [string] $FilterType = "Computer",
    [Parameter(Position=2)]
        [string] $DomainName,
    [Parameter(Position=3)]
        [int] $NumberOfDomainControllersToQuery = 1,
    [Parameter(Position=4)]
        [string] $SpecificDomainControllerToQuery = "*",
    [Parameter(Position=5)]
        [switch] $IntegratedSecurity,
    [Parameter(Position=6)]
        [System.Management.Automation.PSCredential] $Credential,
    [Parameter(Position=7)]
        [switch] $ProgressDisplay,
    [Parameter(Position=8)]
        [switch] $SuppressLogFile,
    [Parameter(Position=9)]
        [string] $LogFilePath,
    [Parameter(Position=10)]
        [switch] $xDTCall,
    [Parameter(Position=11)]
        [string] $DataSource,
    [Parameter(Position=12)]
        [string] $AppVersion
)

#region VARIABLES

#region CONSTANTS
Set-Variable -Name XmlFileExtension -Option Constant -Value 'xml';
Set-Variable -Name LogFileExtension -Option Constant -Value 'log';
Set-Variable -Name DiscoveryDate -Option Constant -Value (Get-Date -format s);
Set-Variable -Name FileDate -Option Constant -Value (Get-Date -format 'M-d-yyyy H.m.s');
Set-Variable -Name Tab -Option Constant -Value ([char]9);
Set-Variable -Name DefaultDate -Option Constant -Value  (Get-Date '1900-01-01' -format s);
Set-Variable -Name PSVersion -Option Constant -Value $PsVersionTable.PSVersion;
Set-Variable -Name dotNetVersion -Option Constant -Value $PsVersionTable.CLRVersion;
#endregion CONSTANTS

#region WORKING VARIABLES
# Set processing variable values
$LoopCnt = 0;
$ResultCount = 0;
$ControllerCount = 0;
$StartDate = Get-Date;
$ExecutionSuccess = $false;

#Initialize Log File Hashtable object
$LogStore = @{};

#Initialize Output objects
$XmlHeader = '<?xml version="1.0" standalone="yes"?>';
$XmlRootOpen = '<Root>';
$XmlRootClose = '</Root>';

#Initialize Versioning object
$Versioning = '' | Select-Object 'DataSource', 'AppVersion', 'ScriptVersion', 'DataOriginSource', 'PrimarySourceTool', 'PrimarySourceToolVersion', 'PSVersion', 'dotNetVersion', 'DiscoveryDate', 'AnonymizationIdentifier', 'AnonymizationCheckValue';
$Versioning.DataSource = $DataSource;
$Versioning.AppVersion = $AppVersion;
$Versioning.PSVersion = $PSVersion;
$Versioning.dotNetVersion = $dotNetVersion;
$Versioning.DiscoveryDate = $DiscoveryDate;

# Create a Regex object to remove invalid XML characters from output value
$invalidXmlCharactersRegex = new-object System.Text.RegularExpressions.Regex("[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u10000-\u10FFFF]");

# If a domain controller name is specified be sure it ends with a wildcard and set $NumberOfDomainControllersToQuery to 0.
If ($SpecificDomainControllerToQuery -ne '*')
{
	If ($SpecificDomainControllerToQuery.EndsWith('*') -ne $TRUE)
	{
		$SpecificDomainControllerToQuery += '*';
	}
	$NumberOfDomainControllersToQuery = 0;
}

#Check file parameters
#Be sure the XML path variable ends with a backslash
If ($XmlFilePath.EndsWith('\') -ne $true)
{
	$XmlFilePath += '\';
}

#Be sure the Log file path is defined and ends with a backslash or set its value equal to the XML path
If(!$LogFilePath)
{
	$LogFilePath = $XmlFilePath;
}
If ($LogFilePath.EndsWith('\') -ne $true)
{
	$LogFilePath += '\';
}
#endregion WORKING VARIABLES
#endregion VARIABLES

#region FUNCTIONS
Function Add-LogEntry()
{
	Param
	(
		$LineValue
	)
	$LogStoreLineCount = ($LogStore.Count + 1);
	$LogStore[$LogStoreLineCount] += $LineValue;
};
#endregion FUNCTIONS

#region PROGRAM MAIN
Try 
{
#region PREPROCESS VALIDATION
# Perform initial validation checks before continuing
	Add-LogEntry -LineValue $('Processing Begin: ' + $(Get-Date -format s).Replace('T',' '));

#Capture the current parameter settings
	Add-LogEntry -LineValue $($Tab+'List of parameter values used for this script execution');
	Add-LogEntry -LineValue $($Tab+$Tab+'XmlFilePath = (' + $XmlFilePath + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'FilterType = (' + $FilterType + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'DomainName = (' + $DomainName + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'NumberOfDomainControllersToQuery = (' + $NumberOfDomainControllersToQuery + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'SpecificDomainControllerToQuery = (' + $SpecificDomainControllerToQuery + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'IntegratedSecurity = (' + $(If($IntegratedSecurity){'On'}Else{'Off'}) + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'Credential = (' + $(If($Credential){'Value Supplied'}Else{'Value Not Supplied'}) + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'ProgressDisplay = (' + $(If($ProgressDisplay){'On'}Else{'Off'}) + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'SuppressLogFile = (' + $(If($SuppressLogFile){'On'}Else{'Off'}) + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'LogFilePath = (' + $LogFilePath + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'xDTCall = (' + $(If($xDTCall){'On'}Else{'Off'}) + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'PSVersion = (' + $PSVersion + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'dotNetVersion = (' + $dotNetVersion + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'DataSource = (' + $DataSource + ')');
	Add-LogEntry -LineValue $($Tab+$Tab+'AppVersion = (' + $AppVersion + ')');

	Try
	{
#FilterType check and specific filter related values set
		Add-LogEntry -LineValue $($Tab+'Validating parameters');
		If (($FilterType -ne 'Computer') -and ($FilterType -ne 'User'))
		{
			$ErrorMessage = 'FilterType parameter only accepts the values, "Computer" or "User".';
		}
#Validate Xml file path is accessible
		ElseIf ($(Test-Path $XmlFilePath) -eq $false)
		{
			$ErrorMessage = 'Could not access specified Xml File Path';
		}
#Validate Log file path is accessible
		ElseIf (!$SuppressLogFile -and ($(Test-Path $LogFilePath) -eq $false))
		{
			$ErrorMessage = 'Could not access specified Log File Path';
		}
# Get Credentials if IntegratedSecurity was not select and script was not called by xDT
		ElseIf (!$IntegratedSecurity -and !$Credential)
		{
# If call was initiated by xDT throw an error as it is responsible for collecting
# Credentials of setting IntegratedSecurity flag
			If ($xDTCall)
			{
				$ErrorMessage = 'Valid User Domain credentials required or IntegratedSecurity must be specified';
			}
			Else
			{
# Otherwise get them from the user
				Add-LogEntry -LineValue $($Tab+'Getting User Domain credentials');
				Try
				{
					If ($psversiontable.psversion.major -lt 3)
					{
						$Credential = Get-Credential;
					}
					Else
					{
						$Credential = Get-Credential -Message 'User Domain credentials';
					}
				}
				Catch
				{
					$ErrorMessage = 'Valid User Domain credentials required or IntegratedSecurity must be specified';
				}
			}
		}
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message;
	}

#An initial cursory check that the user entered at least a name before continuing (did not just hit escape)
#A more thorough check will follow
	If (!$ErrorMessage -and !$IntegratedSecurity)
	{
		Add-LogEntry -LineValue $($Tab+'Validating User Domain credentials entered');
		If (!$Credential.UserName)
		{
			$ErrorMessage = 'Valid credentials required';
		}
	}	
#endregion PREPROCESS VALIDATION

#region COLLECT AND REPORT
#Setting final pre-processing variables
	If (!$ErrorMessage)
	{
		If ($FilterType -eq 'Computer')
		{
			$Versioning.DataOriginSource = 'Active Directory Computers';
			$Versioning.ScriptVersion = '5';
			$Versioning.PrimarySourceTool = 'Active Directory Computers';
			$Versioning.PrimarySourceToolVersion = '5.0.101';
			$FileNamePrefix = 'adcomputers';
			$Filter = '(objectCategory=Computer)';
		}
		Else
		{
			$Versioning.DataOriginSource = 'Active Directory Users';
			$Versioning.ScriptVersion = '5';
			$Versioning.PrimarySourceTool = 'Active Directory Users';
			$Versioning.PrimarySourceToolVersion = '5.0.101';
			$FileNamePrefix = 'adusers';
			$Filter = '(objectCategory=User)';
		}

#Capture the current processing settings
		Add-LogEntry -LineValue $($Tab+$Tab+'List of processing values used for this script execution');
		Add-LogEntry -LineValue $($Tab+$Tab+$Tab+'DataOriginName = (' + $Versioning.DataOriginSource + ')');
		Add-LogEntry -LineValue $($Tab+$Tab+$Tab+'ScriptVersion = (' + $Versioning.ScriptVersion + ')');
		Add-LogEntry -LineValue $($Tab+$Tab+$Tab+'PrimarySourceTool = (' + $Versioning.PrimarySourceTool + ')');
		Add-LogEntry -LineValue $($Tab+$Tab+$Tab+'PrimarySourceToolVersion = (' + $Versioning.PrimarySourceToolVersion + ')');

#Set LogFileName values
		$LogFileName = $LogFilePath + $FileNamePrefix + '_' + $FileDate + '.' + $LogFileExtension;
	
#If we are using credentials seperate the Username and Password
		If ($Credential)
		{
			$UserName = $Credential.username;
			$Password = $Credential.GetNetworkCredential().Password;
		}

		Add-LogEntry -LineValue $($Tab+'Building initial DirectoryContext objects');
# UserName should only exist is credentials where supplied ($IntegratedSecurity = FALSE)
		If ($UserName -ne $null)
		{
# Constructing Domain Context object using supplied Credentials and $DomainName
			If ($DomainName -ne '')
			{
				$DomContext = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $DomainName, $UserName, $Password);
			}
			Else
			{
# Constructing Domain object using supplied Credentials only
				$DomContext = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain',$UserName,$Password);
			}
		}
		ElseIf ($DomainName -ne '')
		{
# Constructing Domain Context using supplied $DomainName
			$DomContext = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $DomainName);
		}
		Else
		{
# Constructing Domain Context for user and domain
			$DomContext = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain');
		}

		Add-LogEntry -LineValue $($Tab+'Connecting to Domain');
		Try
		{
			$D = [System.DirectoryServices.Activedirectory.Domain]::GetDomain($DomContext);
		}
		Catch
		{
			$ErrorMessage = $_.Exception.Message;
			Throw;
		}
		Add-LogEntry -LineValue $($Tab+$Tab+'Connected to Domain - ' + $D.Name);
		$MaxControllerCount = $D.DomainControllers.Count;
		$Domain = [ADSI]$('LDAP://'+$D);

#Opening XmlFileName stream
		Add-LogEntry -LineValue $($Tab+'Opening XML File stream for writing');
		$XmlFileName = $XMLFilePath + $FileNamePrefix + '_' + $FileDate + '.' + $xMLFileExtension;
		Try
		{
			$XMLFileStream = New-Object System.IO.StreamWriter $XmlFileName;
			$XMLFileStream.AutoFlush = $true;
		}
		Catch
		{
			$ErrorMessage = 'Could not create XML File';
			Throw;
		}

# If we are specifying the number of domain comtrollers to check be sure the max number of controllers is more than that number
		If (($NumberOfDomainControllersToQuery -gt 0) -and ($NumberOfDomainControllersToQuery -lt $MaxControllerCount))
		{
			$MaxControllerCount = $NumberOfDomainControllersToQuery;
		}

		Add-LogEntry -LineValue $($Tab+'Creating Xml file');
#Output Header Elements
		Add-LogEntry -LineValue $($Tab+$Tab+'Adding Xml Header element');
		$XMLFileStream.WriteLine($XmlHeader);

#Output Opening Root element
		Add-LogEntry -LineValue $($Tab+$Tab+'Adding opening Root element');
		$XMLFileStream.WriteLine($XmlRootOpen);

#Output Versioning element
		$Table = 'Versioning';
		Add-LogEntry -LineValue $($Tab+$Tab+'Adding Versioning element');
		$Versioning |% {$_.psobject.properties |% {$XMLFileStream.WriteLine($Tab + '<' + $Table + '>')} {$XMLFileStream.WriteLine($($Tab + $Tab + '<' + $_.name + '>' + $_.value + '</' + $_.name + '>'))} {$XMLFileStream.WriteLine($Tab + '</' + $Table + '>')}};

		$Table = 'Object';
		Add-LogEntry -LineValue $($Tab+'Processing Domain Controllers');

# Enumerate through Domain Controllers.
		ForEach ($DC In $D.DomainControllers)
		{
# Initialize processing variables
			$Controller = $DC.Name;
			$LoopCnt++;

# Process only specified Controller if supplied (* = all)
			If ($SpecificDomainControllerToQuery -ne '*')
			{
				If ($Controller -notlike $SpecificDomainControllerToQuery)
				{
					Continue;
				}
			}

# Poll specified number of Domain Controllers (0 = all)
			If ($NumberOfDomainControllersToQuery -ne 0)
			{
				If ($LoopCnt -gt $NumberOfDomainControllersToQuery)
				{
					Break;
				}
			}

			$ControllerCount++;
			If ($SpecificDomainControllerToQuery -ne '*')
			{
				$ControllerDisplayCount = 1;
				$MaxControllerDisplayCount = 1;
			}
			Else
			{
				$ControllerDisplayCount = $ControllerCount;
				$MaxControllerDisplayCount = $MaxControllerCount;
			}

			Add-LogEntry -LineValue $($Tab+$Tab+'Accessing Domain Controller - ' + $Controller);
			If ($ProgressDisplay)
			{
				Write-Progress -Id 0 -Activity $('Collecting '+$FilterType+' Data') -Status $('Processing Controller '+$ControllerDisplayCount+' of '+$MaxControllerDisplayCount) -CurrentOperation $Controller;
			}

# Set DirectorySearcher parameters
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher;
			$Searcher.SearchRoot = "LDAP://$Controller/" + $Domain.distinguishedName;
			$Searcher.PageSize = 1000;
			$Searcher.Filter = $Filter;
			$Searcher.SearchScope = 'Subtree';

# Select and loop through all server on the selected Domain Controller
			Try
			{
				$Results = $Searcher.FindAll();
			}
			Catch
			{
				Add-LogEntry -LineValue $($Tab+$Tab+$Tab+'Domain Controller access problem - skipping Controller');
				Continue;
			}
			$ObjectCount = 0;
			$ObjectCollectStart = Get-Date;
			
			Add-LogEntry -LineValue $($Tab+$Tab+$Tab+'Adding Object elements');
			ForEach ($Result In $Results)
			{
#Initialize variables for capture
				$Store = '' | Select 'DomainController', 'SAMAccountName', 'Name', 'whenchanged', 'whencreated', 'adspath', 'samaccounttype',
					'useraccountcontrol', 'objectcategory', 'LastLogon', 'LastLogonTimeStamp', 'LastPasswordReset', 'CN', 'ServicePrincipalName',
					'MachineRole',	'OperatingSystem', 'OperatingSystemVersion', 'OperatingSystemServicePack';

				$WhenChanged = $null;
				$WhenCreated = $null;
				$LastLogon = $null;
				$LastLogonTimeStamp = $null;
				$LastPasswordReset = $null;
				$ServicePrincipalName = $null;
				$MachineRole = $null;

#Collect and clean values
				$Store.DomainController = $Controller;
				$Store.CN = $Result.Properties.Item('cn');
		
				If ($ProgressDisplay)
				{
					Write-Progress -Id 1 -ParentId 0 -Activity 'Objects Processed' -Status $('Object '+$ObjectCount) -CurrentOperation $Store.CN;
				}

				$LastLogon = $Result.Properties.Item('lastlogon');
				Try
				{
					If ($LastLogon.Count -eq 0)
					{
						$LastLogon = $DefaultDate;
					}
					ElseIf ($LastLogon.Item(0) -eq 0)
					{
						$LastLogon = $DefaultDate;
					}
					Else
					{
						$LastLogon = Get-Date (([DateTime]$LastLogon.Item(0)).AddYears(1600).ToLocalTime()) -format s;
					}
				}
				Catch
				{
					$LastLogon = $DefaultDate;
				}
				$Store.LastLogon = $LastLogon;
				
				$WhenChanged = $Result.Properties.Item('whenchanged');
				Try
				{
					If ($WhenChanged.Count -eq 0)
					{
						$WhenChanged = $DefaultDate;
					}
					ElseIf ($WhenChanged.Item(0) -eq 0)
					{
						$WhenChanged = $DefaultDate;
					}
					Else
					{
						$WhenChanged = Get-Date ($WhenChanged.Item(0)) -format s;
					}
				}
				Catch
				{
					$WhenChanged = $DefaultDate;
				}
				$Store.whenchanged = $WhenChanged;
		
				$WhenCreated = $Result.Properties.Item('whencreated');
				Try
				{
					If ($WhenCreated.Count -eq 0)
					{
						$WhenCreated = $DefaultDate;
					}
					ElseIf ($WhenCreated.Item(0) -eq 0)
					{
						$WhenCreated = $DefaultDate;
					}
					Else
					{
						$WhenCreated = Get-Date ($WhenCreated.Item(0)) -format s;
					}
				}
				Catch
				{
					$WhenCreated = $DefaultDate;
				}
				$Store.whencreated = $WhenCreated;

				$LastLogonTimeStamp = $Result.Properties.Item('lastLogontimeStamp');
				Try
				{
					If ($LastLogonTimeStamp.Count -eq 0)
					{
						$LastLogonTimeStamp = $DefaultDate;
					}
					ElseIf ($LastLogonTimeStamp.Item(0) -eq 0)
					{
						$LastLogonTimeStamp = $DefaultDate;
					}
					Else
					{
						$LastLogonTimeStamp = Get-Date (([DateTime]$LastLogonTimeStamp.Item(0)).AddYears(1600).ToLocalTime()) -format s;
					}
				}
				Catch
				{
					$LastLogonTimeStamp = $DefaultDate;
				}
				$Store.LastLogonTimeStamp = $LastLogonTimeStamp;
				
				$LastPasswordReset = $Result.Properties.Item('pwdlastset');
				Try
				{
					If ($LastPasswordReset.Count -eq 0)
					{
						$LastPasswordReset = $DefaultDate;
					}
					ElseIf ($LastPasswordReset.Item(0) -eq 0)
					{
						$LastPasswordReset = $DefaultDate;
					}
					Else
					{
						$LastPasswordReset = Get-Date (([DateTime]$LastPasswordReset.Item(0)).AddYears(1600).ToLocalTime()) -format s;
					}
				}
				Catch
				{
					$LastPasswordReset = $DefaultDate;
				}
				$Store.LastPasswordReset = $LastPasswordReset;
				
				$ResultEntry = $Result.GetDirectoryEntry();
				Try
				{
					$ServicePrincipalName = $ResultEntry.Properties.Item('servicePrincipalName');
				}
				Catch
				{
					$ServicePrincipalName = '';
				}
				$Store.ServicePrincipalName = $ServicePrincipalName;
				
				Try
				{
					$MachineRole = $ResultEntry.Properties.Item('machinerole');
				}
				Catch
				{
					$MachineRole = '';
				}
				$Store.MachineRole = $MachineRole;
				
				$Store.SAMAccountName = $Result.Properties.Item('SAMAccountName');
				$Store.Name = $Result.Properties.Item('SAMAccountName');
				$Store.adspath = $Result.Properties.Item('adspath');
				$Store.samaccounttype = $Result.Properties.Item('samaccounttype');
				$Store.useraccountcontrol = $Result.Properties.Item('useraccountcontrol');
				$Store.objectcategory = $Result.Properties.Item('objectcategory');
				$Store.OperatingSystem = $Result.Properties.Item('operatingsystem');
				$Store.OperatingSystemVersion = $Result.Properties.Item('operatingsystemversion');
				$Store.OperatingSystemServicePack = $Result.Properties.Item('operatingsystemservicepack');

#Write element to file
				$Store |% {$_.psobject.properties | where {$_.value -ne $null -and $_.value -ne '-'} |% {$XMLFileStream.WriteLine($Tab + '<' + $Table + '>')} {$XMLFileStream.WriteLine($($Tab + $Tab + '<' + $_.name + '>' + $($invalidXmlCharactersRegex.Replace($($_.value), '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace("'", '&apos;').replace('"', '&quot;') + '</' + $_.name + '>')))} {$XMLFileStream.WriteLine($Tab + '</' + $Table + '>')}};
#Increment Row Counter
				$ResultCount++;

#Check for and process any Child rows
				ForEach ($ResultChild In $ResultEntry.Children)
				{
#Initialize variables for capture
					$ChildStore = '' | Select 'DomainController', 'SAMAccountName', 'Name', 'whenchanged', 'whencreated', 'adspath', 'samaccounttype',
					'useraccountcontrol', 'objectcategory', 'LastLogon', 'LastLogonTimeStamp', 'LastPasswordReset', 'CN', 'ChildCN', 'ChildCNKeywords',
					'ServicePrincipalName',	'MachineRole',	'OperatingSystem', 'OperatingSystemVersion', 'OperatingSystemServicePack';

#Copy parent values into Child element
					$ChildStore.DomainController = $Store.DomainController;
					$ChildStore.CN = $Store.CN;
					$ChildStore.LastLogon = $Store.LastLogon;
					$ChildStore.whenchanged = $Store.whenchanged;
					$ChildStore.whencreated = $Store.whencreated;
					$ChildStore.LastLogonTimeStamp = $Store.LastLogonTimeStamp;
					$ChildStore.LastPasswordReset = $Store.LastPasswordReset;
					$ChildStore.ServicePrincipalName = $Store.ServicePrincipalName;
					$ChildStore.MachineRole = $Store.MachineRole;
					$ChildStore.SAMAccountName = $Store.SAMAccountName;
					$ChildStore.Name = $Store.Name;
					$ChildStore.adspath = $Store.adspath;
					$ChildStore.samaccounttype = $Store.samaccounttype;
					$ChildStore.useraccountcontrol = $Store.useraccountcontrol;
					$ChildStore.objectcategory = $Store.objectcategory;
					$ChildStore.OperatingSystem = $Store.OperatingSystem;
					$ChildStore.OperatingSystemVersion = $Store.OperatingSystemVersion;
					$ChildStore.OperatingSystemServicePack = $Store.OperatingSystemServicePack;
#Add Child specific values
					$ChildStore.ChildCN = $ResultChild.cn[0];
					$ChildStore.ChildCNKeywords = $ResultChild.Properties.Item('keywords');
#Write element to file
					$ChildStore |% {$_.psobject.properties | where {$_.value -ne $null -and $_.value -ne '-'} |% {$XMLFileStream.WriteLine($Tab + '<' + $Table + '>')} {$XMLFileStream.WriteLine($($Tab + $Tab + '<' + $_.name + '>' + $($invalidXmlCharactersRegex.Replace($($_.value), '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace("'", '&apos;').replace('"', '&quot;') + '</' + $_.name + '>')))} {$XMLFileStream.WriteLine($Tab + '</' + $Table + '>')}};
#Increment Row Counter
					$ResultCount++;
				}
#Increment Server element counter
				$ObjectCount++;

				If ($ProgressDisplay)
				{
					Write-Progress -Id 1 -ParentId 0 -Activity "Objects Processed " -Status "Object $($ObjectCount)" -CurrentOperation $Store.CN;
				}
			}

			$ObjectCollectEnd = Get-Date;
			$TimeDiff = $ObjectCollectEnd - $ObjectCollectStart;
			$TimeMilliSeconds = [int]$TimeDiff.TotalMilliSeconds;
			Add-LogEntry -LineValue $($Tab+$Tab+$Tab+$Tab+$ObjectCount + ' element entries added in ' + $TimeMilliSeconds + ' milliseconds');
		}

#If a specific controller was requested and failed to return any values
		If ($SpecificDomainControllerToQuery -ne '*' -and $ResultCount -eq 0)
		{
			$ErrorMessage = $('Specified Controller - ' + $SpecificDomainControllerToQuery.Replace('*', '') + ' not found.');
		}
		ElseIf ($ResultCount -eq 0)
		{
			$ErrorMessage = $('No elements Processed using criteria specified.');
		}	

		$EndDate = Get-Date;
		$TimeDiff = $EndDate - $StartDate;
		$TimeMilliSeconds = [int]$TimeDiff.TotalMilliSeconds;

# Write ProcessResults node if there is anything to write
		Add-LogEntry -LineValue $($Tab+$Tab+'Adding ProcessResult element');

		$Table = 'ProcessResult';
		$WriteList = 'ElementsCollected', 'ElementsRequested', 'ElementItemsTotal', 'ProcessTimeStart', 'ProcessTimeEnd', 'ProcessTimeInMilliseconds', 'ExecutionStatus', 'ExecutionStatusMessage';
		$i=@{};
		$i.ElementsCollected = $MaxControllerCount;
		$i.ElementsRequested = $ControllerCount;
		$i.ElementItemsTotal = $ResultCount;
		$i.ProcessTimeStart = $(Get-Date $StartDate -format s);
		$i.ProcessTimeEnd = $(Get-Date $EndDate -format s);
		$i.ProcessTimeInMilliseconds = $TimeMilliSeconds;
		If ($MaxControllerCount -ne $ControllerCount)
		{
			$i.ExecutionStatus = 'Incomplete';
			$i.ExecutionStatusMessage = 'Missing controller elements.';
		}
		Else
		{
			$i.ExecutionStatus = 'Success';
			$i.ExecutionStatusMessage = 'All requested controller elements collected.';
		}
		$ProcessResult = New-Object PSObject -Property $i;
		$ProcessResult | Select-Object $WriteList | ForEach-Object {$_.psobject.properties | ForEach-Object {$XMLFileStream.WriteLine($Tab + '<' + $Table + '>')} {$XMLFileStream.WriteLine($($Tab + $Tab + '<' + $_.name + '>' + $_.value + '</' + $_.name + '>'))} {$XMLFileStream.WriteLine($Tab + '</' + $Table + '>')}};
#Output closing element
		Add-LogEntry -LineValue $($Tab+$Tab+'Adding closing Root element');
		$XMLFileStream.WriteLine($XmlRootClose);
	}
#endregion COLLECT AND REPORT

#region PROCESS SUCCESS FLAG
	If (!$ErrorMessage)
	{
#If we made it all the way to the end without terminating set statue true
		$ExecutionSuccess = $true;
	}
}
Catch
{
	$ErrorMessage = $_.Exception.Message;
}
Finally
{
#Close XML Stream
	Try
	{
		If ($XMLFileStream)
		{
			$XMLFileStream.Close();
			Add-LogEntry -LineValue $($Tab+'Xml file ' + $XmlFilename + ' created.');
		}
		Else
		{
#If there was no file to close, assume no file was opened and clear file name for output
			$XmlFileName = '';
		}
	}
	Catch
	{
		If ($ErrorMessage)
		{
			$ErrorMessage += ': Failed to properly close of XML File.';
		}
		Else
		{
			$ErrorMessage = 'Failed to properly close of XML File';
		}
	}
	
#Create and write Log file if not Suppressed
	If (!$SuppressLogFile)
	{
#If we fell into the PROGRAM MAIN Catch we need to close our processing time stamp
		If (!$EndDate)
		{
			$EndDate = Get-Date;
			$TimeDiff = $EndDate - $StartDate;
			$TimeMilliSeconds = [int]$TimeDiff.TotalMilliSeconds;
		}
#Write log file
		Try
		{
			If (!$LogFileName)
			{
				If ($FilterType)
				{
					$LogFileName = $LogFilePath + 'ad' + $FilterType.ToLower() + '_' + $FileDate + '.' + $LogFileExtension;
				}
				Else
				{
					$LogFileName = $LogFilePath + 'ad_' + $FileDate + '.' + $LogFileExtension;
				}
			}
			$LogFileStream = New-Object System.IO.StreamWriter $LogFileName;
			$LogFileStream.AutoFlush = $true;
			Add-LogEntry -LineValue $($Tab+'Script processing time in milliseconds: ' + $TimeMilliSeconds);
			If (!$ExecutionSuccess)
			{
#If the script is exited with a Ctrl+C the flag will not be set and no error will have been generated
				If (!$ErrorMessage)
				{
					$ErrorMessage = 'Script execution terminated - file write incomplete.';
				}
				Add-LogEntry -LineValue $($Tab+'ExecutionStatus = Failure');
				Add-LogEntry -LineValue $($Tab+'ExecutionStatusMessage = ERROR: ' + $ErrorMessage);
			}
			Else
			{
				If ($MaxControllerCount -ne $ControllerCount)
				{
					Add-LogEntry -LineValue $($Tab+'ExecutionStatus = Incomplete');
					Add-LogEntry -LineValue $($Tab+'ExecutionStatusMessage = Missing controller elements.');
				}
				Else
				{
					Add-LogEntry -LineValue $($Tab+'ExecutionStatus = Success');
					Add-LogEntry -LineValue $($Tab+'ExecutionStatusMessage = All requested controller elements collected.');
				}
			}
			Add-LogEntry -LineValue $('Processing End: ' + $((Get-Date -format s).Replace('T',' ')));

			$LogStore.GetEnumerator() | Sort-Object Name | ForEach-Object {$LogFileStream.WriteLine($_.Value)} -ErrorAction SilentlyContinue;
			$LogFileStream.Close();
		}
		Catch
		{
			If (!$xDTCall)
			{
				If ($ErrorMessage)
				{
					$ErrorMessage += ': Could not create LOG File';
				}
				Else
				{
					$ErrorMessage = $('Could not create LOG File');
				}
			}
		}
	}
	$CollectionResults = '' | Select-Object 'CollectionSuccess', 'FileName', 'Error';
	$CollectionResults.CollectionSuccess = $ExecutionSuccess;
	$CollectionResults.FileName = $XmlFileName;
	$CollectionResults.Error = $ErrorMessage;
}
#endregion PROGRAM MAIN
If ($xDTCall)
{
	$CollectionResults;
}
Else
{
	$CollectionResults | fl;
}

# SIG # Begin signature block
# MIIkFAYJKoZIhvcNAQcCoIIkBTCCJAECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDLDQx4eLVSpvYb
# M/f2Lc5tAZS+TqHtI36QY9Vlccm22qCCDZMwggYRMIID+aADAgECAhMzAAAAjoeR
# pFcaX8o+AAAAAACOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMTYxMTE3MjIwOTIxWhcNMTgwMjE3MjIwOTIxWjCBgzEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjENMAsGA1UECxMETU9Q
# UjEeMBwGA1UEAxMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEA0IfUQit+ndnGetSiw+MVktJTnZUXyVI2+lS/qxCv
# 6cnnzCZTw8Jzv23WAOUA3OlqZzQw9hYXtAGllXyLuaQs5os7efYjDHmP81LfQAEc
# wsYDnetZz3Pp2HE5m/DOJVkt0slbCu9+1jIOXXQSBOyeBFOmawJn+E1Zi3fgKyHg
# 78CkRRLPA3sDxjnD1CLcVVx3Qv+csuVVZ2i6LXZqf2ZTR9VHCsw43o17lxl9gtAm
# +KWO5aHwXmQQ5PnrJ8by4AjQDfJnwNjyL/uJ2hX5rg8+AJcH0Qs+cNR3q3J4QZgH
# uBfMorFf7L3zUGej15Tw0otVj1OmlZPmsmbPyTdo5GPHzwIDAQABo4IBgDCCAXww
# HwYDVR0lBBgwFgYKKwYBBAGCN0wIAQYIKwYBBQUHAwMwHQYDVR0OBBYEFKvI1u2y
# FdKqjvHM7Ww490VK0Iq7MFIGA1UdEQRLMEmkRzBFMQ0wCwYDVQQLEwRNT1BSMTQw
# MgYDVQQFEysyMzAwMTIrYjA1MGM2ZTctNzY0MS00NDFmLWJjNGEtNDM0ODFlNDE1
# ZDA4MB8GA1UdIwQYMBaAFEhuZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0Nv
# ZFNpZ1BDQTIwMTFfMjAxMS0wNy0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsG
# AQUFBzAChkVodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01p
# Y0NvZFNpZ1BDQTIwMTFfMjAxMS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkq
# hkiG9w0BAQsFAAOCAgEARIkCrGlT88S2u9SMYFPnymyoSWlmvqWaQZk62J3SVwJR
# avq/m5bbpiZ9CVbo3O0ldXqlR1KoHksWU/PuD5rDBJUpwYKEpFYx/KCKkZW1v1rO
# qQEfZEah5srx13R7v5IIUV58MwJeUTub5dguXwJMCZwaQ9px7eTZ56LadCwXreUM
# tRj1VAnUvhxzzSB7pPrI29jbOq76kMWjvZVlrkYtVylY1pLwbNpj8Y8zon44dl7d
# 8zXtrJo7YoHQThl8SHywC484zC281TllqZXBA+KSybmr0lcKqtxSCy5WJ6PimJdX
# jrypWW4kko6C4glzgtk1g8yff9EEjoi44pqDWLDUmuYx+pRHjn2m4k5589jTajMW
# UHDxQruYCen/zJVVWwi/klKoCMTx6PH/QNf5mjad/bqQhdJVPlCtRh/vJQy4njpI
# BGPveJiiXQMNAtjcIKvmVrXe7xZmw9dVgh5PgnjJnlQaEGC3F6tAE5GusBnBmjOd
# 7jJyzWXMT0aYLQ9RYB58+/7b6Ad5B/ehMzj+CZrbj3u2Or2FhrjMvH0BMLd7Hald
# G73MTRf3bkcz1UDfasouUbi1uc/DBNM75ePpEIzrp7repC4zaikvFErqHsEiODUF
# he/CBAANa8HYlhRIFa9+UrC4YMRStUqCt4UqAEkqJoMnWkHevdVmSbwLnHhwCbww
# ggd6MIIFYqADAgECAgphDpDSAAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYD
# VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
# MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3Nv
# ZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5
# MDlaFw0yNjA3MDgyMTA5MDlaMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIw
# MTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQ
# TTS68rZYIZ9CGypr6VpQqrgGOBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULT
# iQ15ZId+lGAkbK+eSZzpaF7S35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYS
# L+erCFDPs0S3XdjELgN1q2jzy23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494H
# DdVceaVJKecNvqATd76UPe/74ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZ
# PrGMXeiJT4Qa8qEvWeSQOy2uM1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5
# bmR/U7qcD60ZI4TL9LoDho33X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGS
# rhwjp6lm7GEfauEoSZ1fiOIlXdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADh
# vKwCgl/bwBWzvRvUVUvnOaEP6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON
# 7E1JMKerjt/sW5+v/N2wZuLBl4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xc
# v3coKPHtbcMojyyPQDdPweGFRInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqw
# iBfenk70lrC8RqBsmNLg1oiMCwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMC
# AQAwHQYDVR0OBBYEFEhuZOVQBdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQM
# HgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1Ud
# IwQYMBaAFHItOgIxkEO5FAVO4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0
# dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0Nl
# ckF1dDIwMTFfMjAxMV8wM18yMi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUF
# BzAChkJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0Nl
# ckF1dDIwMTFfMjAxMV8wM18yMi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGC
# Ny4DMIGDMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2RvY3MvcHJpbWFyeWNwcy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcA
# YQBsAF8AcABvAGwAaQBjAHkAXwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZI
# hvcNAQELBQADggIBAGfyhqWY4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4s
# PvjDctFtg/6+P+gKyju/R6mj82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKL
# UtCw/WvjPgcuKZvmPRul1LUdd5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7
# pKkFDJvtaPpoLpWgKj8qa1hJYx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft
# 0N3zDq+ZKJeYTQ49C/IIidYfwzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4
# MnEnGn+x9Cf43iw6IGmYslmJaG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxv
# FX1Fp3blQCplo8NdUmKGwx1jNpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG
# 0QaxdR8UvmFhtfDcxhsEvt9Bxw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf
# 0AApxbGbpT9Fdx41xtKiop96eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkY
# S//WsyNodeav+vyL6wuA6mk7r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrv
# QQqxP/uozKRdwaGIm1dxVk5IRcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIV
# 1zCCFdMCAQEwgZUwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAA
# AI6HkaRXGl/KPgAAAAAAjjANBglghkgBZQMEAgEFAKCBxjAZBgkqhkiG9w0BCQMx
# DAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkq
# hkiG9w0BCQQxIgQgx575JBrxUhZxlHT4xp6rPTVA4DlMXjnsyUyXcvBX8k0wWgYK
# KwYBBAGCNwIBDDFMMEqgMIAuAFcAbwByAGsAUwBwAGEAYwBlACAAQwBvAG0AcABh
# AG4AaQBvAG4AIABBAHAAcKEWgBRodHRwOi8vbXN3b3Jrc3BhY2UvIDANBgkqhkiG
# 9w0BAQEFAASCAQCzBwYPVCI/Wbx/wKqfwfPFiWyTRpSM0/FNPhE9ZwmHr+J+SwTT
# QnonDF+J2+EC3pRoUc/I+VDrIHfbdyGJHkb+oNOhPDx4wFXcSqhgjS0bsmOr2j6B
# YvDLzTslNjYR94SOn8Tr/0yNottcwvbceJ74SEYaIHVCBFFTNBOS9kZczLlohvS7
# KrhawFJgmPFnimVn6IR2F6X/JaKQqV5ZDyWANzEpQH/wIWaPMwbYFAvPsD+x3fL7
# q9wxMVsUNLX32E2Cc25ScpQmOp6FXxfRec9z49vzu2rcKFUlLUkAyXB4ZHhtbOIx
# NgmGeH5IX7wuAav0+RGpuJcaotHd/zQsHe+yoYITSTCCE0UGCisGAQQBgjcDAwEx
# ghM1MIITMQYJKoZIhvcNAQcCoIITIjCCEx4CAQMxDzANBglghkgBZQMEAgEFADCC
# AT0GCyqGSIb3DQEJEAEEoIIBLASCASgwggEkAgEBBgorBgEEAYRZCgMBMDEwDQYJ
# YIZIAWUDBAIBBQAEIGnkbvnhiZBhC3FFP70WAYC1JD1/7R4WLBfXXqnWjDlXAgZZ
# VDtpT4AYEzIwMTcwNjI5MTgyMDMyLjc5NVowBwIBAYACAfSggbmkgbYwgbMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIx
# JzAlBgNVBAsTHm5DaXBoZXIgRFNFIEVTTjo5OEZELUM2MUUtRTY0MTElMCMGA1UE
# AxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCCDswwggZxMIIEWaADAgEC
# AgphCYEqAAAAAAACMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0
# aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0xMDA3MDEyMTM2NTVaFw0yNTA3MDEy
# MTQ2NTVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqR0NvHcRijog7PwTl/X6f2mUa3RUENWlCgCC
# hfvtfGhLLF/Fw+Vhwna3PmYrW/AVUycEMR9BGxqVHc4JE458YTBZsTBED/FgiIRU
# QwzXTbg4CLNC3ZOs1nMwVyaCo0UN0Or1R4HNvyRgMlhgRvJYR4YyhB50YWeRX4FU
# sc+TTJLBxKZd0WETbijGGvmGgLvfYfxGwScdJGcSchohiq9LZIlQYrFd/XcfPfBX
# day9ikJNQFHRD5wGPmd/9WbAA5ZEfu/QS/1u5ZrKsajyeioKMfDaTgaRtogINeh4
# HLDpmc085y9Euqf03GS9pAHBIAmTeM38vMDJRF1eFpwBBU8iTQIDAQABo4IB5jCC
# AeIwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFNVjOlyKMZDzQ3t8RhvFM2ha
# hW1VMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNV
# HRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYG
# A1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3Js
# L3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBaBggrBgEFBQcB
# AQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kv
# Y2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0MIGgBgNVHSABAf8EgZUw
# gZIwgY8GCSsGAQQBgjcuAzCBgTA9BggrBgEFBQcCARYxaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL1BLSS9kb2NzL0NQUy9kZWZhdWx0Lmh0bTBABggrBgEFBQcCAjA0
# HjIgHQBMAGUAZwBhAGwAXwBQAG8AbABpAGMAeQBfAFMAdABhAHQAZQBtAGUAbgB0
# AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAB+aIUQ3ixuCYP4FxAz2do6Ehb7Prpsz1
# Mb7PBeKp/vpXbRkws8LFZslq3/Xn8Hi9x6ieJeP5vO1rVFcIK1GCRBL7uVOMzPRg
# Eop2zEBAQZvcXBf/XPleFzWYJFZLdO9CEMivv3/Gf/I3fVo/HPKZeUqRUgCvOA8X
# 9S95gWXZqbVr5MfO9sp6AG9LMEQkIjzP7QOllo9ZKby2/QThcJ8ySif9Va8v/rbl
# jjO7Yl+a21dA6fHOmWaQjP9qYn/dxUoLkSbiOewZSnFjnXshbcOco6I8+n99lmqQ
# eKZt0uGc+R38ONiU9MalCpaGpL2eGq4EQoO4tYCbIjggtSXlZOz39L9+Y1klD3ou
# OVd2onGqBooPiRa6YacRy5rYDkeagMXQzafQ732D8OE7cQnfXXSYIghh2rBQHm+9
# 8eEA3+cxB6STOvdlR3jo+KhIq/fecn5ha293qYHLpwmsObvsxsvYgrRyzR30uIUB
# HoD7G4kqVDmyW9rIDVWZeodzOwjmmC3qjeAzLhIp9cAvVCch98isTtoouLGp25ay
# p0Kiyc8ZQU3ghvkqmqMRZjDTu3QyS99je/WZii8bxyGvWbWu3EQ8l1Bx16HSxVXj
# ad5XwdHeMMD9zOZN+w2/XU/pnR4ZOC+8z1gFLu8NoFA12u8JJxzVs341Hgi62jbb
# 01+P3nSISRIwggTaMIIDwqADAgECAhMzAAAAnSCcVndV1CiaAAAAAACdMA0GCSqG
# SIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTE2MDkw
# NzE3NTY0MVoXDTE4MDkwNzE3NTY0MVowgbMxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIxJzAlBgNVBAsTHm5DaXBoZXIg
# RFNFIEVTTjo5OEZELUM2MUUtRTY0MTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgU2VydmljZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANJE
# mJwRWioaLqqfU11tXby2WXaRwCZbA+bIbF+jKutMAEZ0OBS/KnhdsCNM7G5gSOxJ
# 5Ft1pnD989SuVW6OvQQfZz0Z/TFygpShc7EuvPAc1NvvIbjGqbTGwkYHLpnMPiEL
# wy5I3wxqdcU1jtdZnKs7SH6esuD8VJbeE0c5QtBu1kv9vwyk8Avl+ujIiIvunPt1
# 4cRL6MsOZM5X3mCoekrOZRy4ZZYjYjt/BU9ZZt3pDdX4fL7ATN57CpYbzFU5BG8G
# CEE4u/UZ37V6BHcFHOLsjMfxsZpeR27Msh6j2pZ4ge7wB5iAUb66ChQefp46WSSh
# V3MM/kFETpbCVFEPqcUCAwEAAaOCARswggEXMB0GA1UdDgQWBBS8hgjKW2payuS9
# zMuCtBVI6ofloTAfBgNVHSMEGDAWgBTVYzpcijGQ80N7fEYbxTNoWoVtVTBWBgNV
# HR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9NaWNUaW1TdGFQQ0FfMjAxMC0wNy0wMS5jcmwwWgYIKwYBBQUHAQEE
# TjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY1RpbVN0YVBDQV8yMDEwLTA3LTAxLmNydDAMBgNVHRMBAf8EAjAAMBMG
# A1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEBCwUAA4IBAQB/3iQhvVnvtNaL
# ccpZkb4uqEaCu4/fZB195ioLvChnS/75d7+19E6k/ehKDz5nIrNWiW2XCFrsIxT1
# eSoTV4ySF50GIerzqOobO9zbhJpL93IV9p+PJ6j/peLWIImVTUCpFWBeuZcB1zAL
# /0Jqa1bZ7FpcNgOAzBYtasG5M2RP215rf9hvwK6BpTjtOs5dchqMTBXLX5OMst2q
# AC3j/WQoqam+EB3+Fdwnjx+OpAPqjjfbBCVTH+Eyevc7IpDM3CoNwV6GCdU+Vu+r
# JaB6yzJAWPa9CVu2yf97R3l0hqWGndgiDVde4agNxiZOAvb9OvYBrPeXvLmRDmHb
# ndPvpjZpoYIDdTCCAl0CAQEwgeOhgbmkgbYwgbMxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIxJzAlBgNVBAsTHm5DaXBo
# ZXIgRFNFIEVTTjo5OEZELUM2MUUtRTY0MTElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZaIlCgEBMAkGBSsOAwIaBQADFQAYDayzjGgws/h0GbJ4
# zoArNS8I+qCBwjCBv6SBvDCBuTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBOVFMgRVNO
# OjU3RjYtQzFFMC01NTRDMSswKQYDVQQDEyJNaWNyb3NvZnQgVGltZSBTb3VyY2Ug
# TWFzdGVyIENsb2NrMA0GCSqGSIb3DQEBBQUAAgUA3P+v/jAiGA8yMDE3MDYyOTE2
# NTczNFoYDzIwMTcwNjMwMTY1NzM0WjBzMDkGCisGAQQBhFkKBAExKzApMAoCBQDc
# /6/+AgEAMAYCAQACASwwBwIBAAICGtswCgIFAN0BAX4CAQAwNgYKKwYBBAGEWQoE
# AjEoMCYwDAYKKwYBBAGEWQoDAaAKMAgCAQACAxbjYKEKMAgCAQACAwehIDANBgkq
# hkiG9w0BAQUFAAOCAQEAUErLZMo9zeqdny9PMyluqrAkkTvlT5ZIC51hoOLpDgHq
# lcK9TNy7p/dW2T1hYYTxoz+zbot2UGrwFKIrR1UcYCHbi8PNUgYPvhmPsjXYjSM1
# bagS7F7Bi5WDxBsJg6nBL90KCR6O5cYYpnUgg5sGkhSMqLtHyMUMnB63xUs4+Fnz
# NNT3hSHvxhOunlY2luoVJeJyjKSIMEBGsKjAuBV/RvogoUvo4/KJhM1MFJSckJJl
# IlehRhLE9HYRjf7FiO4X6YeSPxyIWOz6LjaLvXPvnEtWD0daj9G9u9+t5EOw64xc
# 0WLFnJ5YS8MwXAWwTfhPdhhkRzxxDST+8H4OQhLt0zGCAvUwggLxAgEBMIGTMHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAAnSCcVndV1CiaAAAAAACd
# MA0GCWCGSAFlAwQCAQUAoIIBMjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# LwYJKoZIhvcNAQkEMSIEIB3e40TWdJA5Ob5Zu/dm3NskTrl6bsNmBLQNnKAnQl+N
# MIHiBgsqhkiG9w0BCRACDDGB0jCBzzCBzDCBsQQUGA2ss4xoMLP4dBmyeM6AKzUv
# CPowgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAJ0g
# nFZ3VdQomgAAAAAAnTAWBBTL8Pw1HiqK0OlIkYmCPdLPPCiI+DANBgkqhkiG9w0B
# AQsFAASCAQCwbqHFxdYahIyrcJQwjY9HuwdPV79j7xuYKwuQY7zal58tXERul6dw
# 0G/9YquQcMTvU0ce7jI4mJpIFcJT8ek/+BHG6oi/N7VKw3pz9hY/AhD8tWvn4v8O
# wfhB7ieu0neceSVR8VkfyAU5be8r12LAzsPymHhFcSAI9DU/T7ZnAiR69LzUI4bL
# sJF4b0kIFmOfzydbp33SFEWpiWWluvSzpfmbhOX6RFZeptBLXYQAr9NDyukcMsxO
# trYJKSiSvxroeds7C/ZJnTul0dw3vYFu2VWOTYl2qFSHJhWHuQZVWcvKuWhmKEkx
# sJ5Py+vYp9gRv2CPbjXR0awGXVTiQ3GU
# SIG # End signature block
