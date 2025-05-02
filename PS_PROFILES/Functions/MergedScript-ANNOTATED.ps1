## Begin ConnectTo-Database
Function Open-Database {
    # Begin Connection Config
    $SQLServer = "<serverName>"
    # Removed unused variable $SQLDB
    # Credentials for sa user
    $CredsPath = ".\Creds\SQLCredsSA.xml"
    if (Test-Path -Path $CredsPath) {
        $SQLCred = Import-Clixml -Path $CredsPath
    } else {
        Write-Error "Credential file not found at $CredsPath. Please ensure the file exists."
        return
    }
    # End Connection Config

    # Begin Region Connection
    Import-Module dbatools
    Write-Host "Building connection string" -ForegroundColor Black -BackgroundColor White
    try {
        Connect-DbaInstance -SqlInstance $SQLServer -Credential $SQLCred -ErrorAction Stop
    } catch {
        Write-Error "Failed to connect to SQL Server instance '$SQLServer'. Error: $_"
        return
    }
    Write-Host "Opening connection to $($SqlServer)"
    # End Region Connection

}
## End ConnectTo-Database
## Begin Copy-Folder
Function Copy-Folder([string]$source, [string]$destination, [bool]$recursive) {
    if (!$(Test-Path($destination))) {
        New-Item $destination -type directory -Force
    }
    ####################################################################################################
    # This Function copies a folder (and optionally, its subfolders)
    #
    # When copying subfolders it calls itself recursively
    #
    # Initializes WebClient object $webClient if not already defined
    #
    # Parameters:
    #   $source      - The url of folder to copy, with trailing /, e.g. http://website/folder/structure/
    #   $destination - The folder to copy $source to, with trailing \ e.g. D:\CopyOfStructure\
    #   $recursive   - True if subfolders of $source are also to be copied or False to ignore subfolders
    #   Return       - None
    ####################################################################################################
    # Initialize WebClient
    if (-not $webClient) {
        $webClient = New-Object System.Net.WebClient
    }

    # Get the file list from the web page
    $webString = $webClient.DownloadString($source)
    $lines = [Regex]::Split($webString, "<br>")
    # Parse each line, looking for files and folders
    foreach ($line in $lines) {
        if ($line.ToUpper().Contains("HREF")) {
            # File or Folder
            if (!$line.ToUpper().Contains("[TO PARENT DIRECTORY]")) {
                # Not Parent Folder entry
                $items = [Regex]::Split($line, """")
                $items = [Regex]::Split($items[2], "(>|<)")
                $item = $items[2]
                if ($line.ToLower().Contains("&lt;dir&gt")) {
                    # Folder
                    if ($recursive) {
                        # Subfolder copy required
                        Copy-Folder "$source$item/" "$destination$item/" $recursive
                    }
                    else {
                        # Subfolder copy not required
                    }
                }
                else {
                    # File
                    $webClient.DownloadFile("$source$item", "$destination$item")
                }
            }
        }
    }
}
## End Copy-Folder
## Begin New-HTMLTable
Function New-HTMLTable {
    param([array]$Array)

    if (-not $Array -or $Array.Count -eq 0) {
        Write-Warning "Input array is empty or null. Returning an empty HTML table."
        return "<table></table>"
    }

    # Convert array to HTML and remove unwanted tags
    $arrHTML = $Array | ConvertTo-Html -Fragment
    $arrHTML[-1] = $arrHTML[-1].ToString().Replace("</body></html>", "")

    # Ensure we return a valid range
    $startIndex = [Math]::Min(5, $arrHTML.Count - 1)
    $endIndex = [Math]::Min(2000, $arrHTML.Count - 1)

    return $arrHTML[$startIndex..$endIndex] -join "`r`n"
}
## End New-HTMLTable
## Begin New-WMIFilters
Function New-WMIFilters {
    Import-Module ActiveDirectory -ErrorAction Stop  # Ensure AD module is loaded

    # Define WMI Filters (Name, Query, Description)
    $WMIFilters = @(
        ('Virtual Machines', 'SELECT * FROM Win32_ComputerSystem WHERE Model = "Virtual Machine"', 'Hyper-V'),
        ('Workstation 32-bit', 'Select * from WIN32_OperatingSystem where ProductType=1 AND AddressWidth = "32"', ''),
        ('Workstation 64-bit', 'Select * from WIN32_OperatingSystem where ProductType=1 AND AddressWidth = "64"', ''),
        ('Workstations', 'SELECT * FROM Win32_OperatingSystem WHERE ProductType = "1"', ''),
        ('Domain Controllers', 'SELECT * FROM Win32_OperatingSystem WHERE ProductType = "2"', ''),
        ('Servers', 'SELECT * FROM Win32_OperatingSystem WHERE ProductType = "3"', ''),
        ('Hotfix', 'SELECT * FROM Win32_QuickFixEngineering WHERE HotFixID = "q147222"', 'Apply policy on computers with specific hotfix'),
        ('Manufacturer Dell', 'SELECT * FROM Win32_ComputerSystem WHERE Manufacturer = "DELL"', '')
    )

    # Get AD Naming Contexts
    try {
        $defaultNamingContext = (Get-ADRootDSE).defaultNamingContext
        $configurationNamingContext = (Get-ADRootDSE).configurationNamingContext
        $domainName = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
    } catch {
        Write-Error "Failed to retrieve AD naming contexts. Ensure you have domain access."
        return
    }

    $msWMIAuthor = "Administrator@$domainName"

    # Process each WMI filter
    foreach ($filter in $WMIFilters) {
        $WMIGUID = [System.Guid]::NewGuid().ToString("B")  # Generate unique GUID
        $WMIDN = "CN=$WMIGUID,CN=SOM,CN=WMIPolicy,CN=System,$defaultNamingContext"

        # Generate timestamp
        $now = Get-Date
        $msWMICreationDate = "{0:yyyyMMddHHmmss.000000-000}" -f $now.ToUniversalTime()

        # Construct attributes
        $Attr = @{
            "msWMI-Name" = $filter[0]
            "msWMI-Parm1" = $filter[2]
            "msWMI-Parm2" = "1;3;10;$($filter[1].Length);WQL;root\CIMv2;$($filter[1]);"
            "msWMI-Author" = $msWMIAuthor
            "msWMI-ID" = $WMIGUID
            "instanceType" = 4
            "showInAdvancedViewOnly" = "TRUE"
            "distinguishedname" = $WMIDN
            "msWMI-ChangeDate" = $msWMICreationDate
            "msWMI-CreationDate" = $msWMICreationDate
        }

        $WMIPath = "CN=SOM,CN=WMIPolicy,CN=System,$defaultNamingContext"

        # Create AD Object
        try {
            New-ADObject -Name $WMIGUID -Type "msWMI-Som" -Path $WMIPath -OtherAttributes $Attr
            Write-Host "Successfully created WMI filter: $($filter[0])"
        } catch {
            Write-Error "Failed to create WMI filter: $($filter[0]). Ensure AD permissions."
        }
    }
}
## End New-WMIFilters
## Begin Expand-ScriptAlias
Function Expand-ScriptAlias {
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
    PROCESS {
        FOREACH ($File in $Path) {
            Write-Verbose -Message '[PROCESS] $File'
			
            TRY {
                # Retrieve file content
                $ScriptContent = (Get-Content $File -Delimiter $([char]0))
				
                # AST Parsing
                $AbstractSyntaxTree = [System.Management.Automation.Language.Parser]::
                ParseInput($ScriptContent, [ref]$null, [ref]$null)
				
                # Find Aliases
                $Aliases = $AbstractSyntaxTree.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true) |
                ForEach-Object -Process {
                    $Command = $_.CommandElements[0]
                    if ($Alias = Get-Alias | Where-Object { $_.Name -eq $Command }) {
						
                        # Output information
                        [PSCustomObject]@{
                            File              = $File
                            Alias             = $Alias.Name
                            Definition        = $Alias.Definition
                            StartLineNumber   = $Command.Extent.StartLineNumber
                            EndLineNumber     = $Command.Extent.EndLineNumber
                            StartColumnNumber = $Command.Extent.StartColumnNumber
                            EndColumnNumber   = $Command.Extent.EndColumnNumber
                            StartOffset       = $Command.Extent.StartOffset
                            EndOffset         = $Command.Extent.EndOffset
							
                        }#[PSCustomObject]
                    }#if ($Alias)
                } | Sort-Object -Property EndOffset -Descending
				
                # The sort-object is important, we change the values from the end first to not lose the positions of every aliases.
                Foreach ($Alias in $Aliases) {
                    # whatif and confirm support
                    if ($psCmdlet.ShouldProcess($file, "Expand Alias: $($Alias.alias) to $($Alias.definition) (startoffset: $($alias.StartOffset))")) {
                        # Remove alias and insert full cmldet name
                        $ScriptContent = $ScriptContent.Remove($Alias.StartOffset, ($Alias.EndOffset - $Alias.StartOffset)).Insert($Alias.StartOffset, $Alias.Definition)
                        # Apply to the file
                        Set-Content -Path $File -Value $ScriptContent -Confirm:$false
                    }
                }#ForEach Alias in Aliases
				
            }#TRY
            CATCH {
                Write-Error -Message $($Error[0].Exception.Message)
            }
        }#FOREACH File in Path
    }#PROCESS
}
## End Expand-ScriptAlias
## Begin Get-ActivationStatus - Needs to be fixed!
Function Get-ActivationStatus {

    $Servers = Get-Content C:\LazyWinAdmin\Servers\Servers-All-Alive2.txt

    ForEach ($server in $Servers) {
        try {
            $wpa = Get-CimInstance SoftwareLicensingProduct -ComputerName $server -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" -Property LicenseStatus -ErrorAction SilentlyContinue
        } catch {
            $status = New-Object ComponentModel.Win32Exception ($_.Exception.ErrorCode)
            $wpa = $null    
        }
    }
    $out = New-Object psobject -Property @{
        ComputerName = $server;
        Status       = [string]::Empty;
    }
    if ($wpa) {
        :outer foreach ($item in $wpa) {
            switch ($item.LicenseStatus) {
                0 { $out.Status = "Unlicensed" }
                1 { $out.Status = "Licensed"; break outer }
                2 { $out.Status = "Out-Of-Box Grace Period"; break outer }
                3 { $out.Status = "Out-Of-Tolerance Grace Period"; break outer }
                4 { $out.Status = "Non-Genuine Grace Period"; break outer }
                5 { $out.Status = "Notification"; break outer }
                6 { $out.Status = "Extended Grace"; break outer }
                default { $out.Status = "Unknown value" }
            }
        }
    }
    else { $out.Status = $status.Message }
    $out
}
## End Get-ActivationStatus - Needs to be fixed!
## Begin Get-AsciiReaction
Function Get-AsciiReaction {
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
    Function Write-Ascii {
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
	
    Switch ($Name) {
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
        default {
            [PSCustomObject][ordered]@{
                'Shrug'       = [char[]]@(175, 92, 95, 40, 12484, 41, 95, 47, 175) -join '' | Write-Ascii
                'Disapproval' = [char[]]@(3232, 95, 3232) -join '' | Write-Ascii
                'TableFlip'   = [char[]]@(40, 9583, 176, 9633, 176, 65289, 9583, 65077, 32, 9531, 9473, 9531, 41) -join '' | Write-Ascii
                'TableBack'   = [char[]]@(9516, 9472, 9472, 9516, 32, 175, 92, 95, 40, 12484, 41) -join '' | Write-Ascii 
                'TableFlip2'  = [char[]]@(9531, 9473, 9531, 32, 65077, 12541, 40, 96, 1044, 180, 41, 65417, 65077, 32, 9531, 9473, 9531) -join '' | Write-Ascii 
                'TableBack2'  = [char[]]@(9516, 9472, 9516, 12494, 40, 32, 186, 32, 95, 32, 186, 12494, 41) -join '' | Write-Ascii 
                'TableFlip3'  = [char[]]@(40, 12494, 3232, 30410, 3232, 41, 12494, 24417, 9531, 9473, 9531) -join '' | Write-Ascii 
                'Denko'       = [char[]]@(40, 180, 65381, 969, 65381, 96, 41) -join '' | Write-Ascii 
                'BlowKiss'    = [char[]]@(40, 42, 94, 51, 94, 41, 47, 126, 9734) -join '' | Write-Ascii 
                'Lenny'       = [char[]]@(40, 32, 865, 176, 32, 860, 662, 32, 865, 176, 41) -join '' | Write-Ascii 
                'Angry'       = [char[]]@(40, 65283, 65439, 1044, 65439, 41) -join '' | Write-Ascii 
                'DontKnow'    = [char[]]@(9488, 40, 39, 65374, 39, 65307, 41, 9484) -join '' | Write-Ascii 
            }
        }
    }
}
## End Get-AsciiReaction
## Begin Get-ByOwner
Function Get-ByOwner {
    param([string]$OwnerMatch)
    Get-ChildItem -Recurse C:\ -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Get-ACL $_.FullName -ErrorAction Stop
        } catch {
            Write-Verbose "Skipping inaccessible item: $($_.FullName)"
            continue
        }
    } | Where-Object { $_.Owner -Match $OwnerMatch }
}
## End Get-ByOwner
## Begin Get-CDROMDetails
Function Get-CDROMDetails {                        
    [cmdletbinding()]                        
    param(                        
        [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]                        
        [string[]]$ComputerName = $env:COMPUTERNAME                        
    )                        
            
    begin {}                        
    process {                        
        foreach ($Computer in $COmputerName) {                        
            $object = New-Object �TypeName PSObject �Prop(@{                        
                    'ComputerName' = $Computer.ToUpper();                        
                    'CDROMDrive'   = $null;                        
                    'Manufacturer' = $null                        
                })                        
            if (!(Test-Connection -ComputerName $Computer -Count 1 -ErrorAction 0)) {                        
                Write-Verbose "$Computer is OFFLINE"                        
            }                        
            try {                        
                $cd = Get-WMIObject -Class Win32_CDROMDrive -ComputerName $Computer -ErrorAction Stop                        
            }
            catch {                        
                Write-Verbose "Failed to Query WMI Class"                        
                Continue;                        
            }                        
            
            $Object.CDROMDrive = $cd.Drive                        
            $Object.Manufacturer = $cd.caption                        
            $Object                           
            
        }                        
    }                        
    ## End Get-CDROMDetails
            
    end {}               
}
## End Get-CDROMDetails
## Begin Get-ComputerHardwareSpecification
Function Get-ComputerHardwareSpecification {
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
Function Get-ComputerInfo {

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
        [Parameter(ValueFromPipeline = $true)]
        [String[]]$ComputerName = "LocalHost",

        [String]$ErrorLog = ".\Errors.log",

        [Alias("RunAs")]
        [System.Management.Automation.PSCredential]
        $Credential
    )#PARAM

    BEGIN {}#PROCESS BEGIN

    PROCESS {
        FOREACH ($Computer in $ComputerName) {
            Write-Verbose -Message "PROCESS - Querying $Computer ..."
                
            TRY {
                $Splatting = @{
                    ComputerName = $Computer
                }

                IF ($PSBoundParameters["Credential"]) {
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
                FOREACH ($Proc in $Processors) {
                    IF ($null -eq $Proc.numberofcores) {
                        IF ($null -ne $Proc.SocketDesignation) { $Sockets++ }
                        $Cores++
                    }
                    ELSE {
                        $Sockets++
                        $Cores += $proc.numberofcores
                    }#ELSE
                }#FOREACH $Proc in $Processors

            }
            CATCH {
                $Everything_is_OK = $false
                Write-Warning -Message "Error on $Computer"
                $Computer | Out-file -FilePath $ErrorLog -Append -ErrorAction Continue
                $ProcessError | Out-file -FilePath $ErrorLog -Append -ErrorAction Continue
                Write-Warning -Message "Logged in $ErrorLog"

            }#CATCH


            IF ($Everything_is_OK) {
                Write-Verbose -Message "PROCESS - $Computer - Building the Output Information"
                $Info = [ordered]@{
                    "ComputerName"       = $OperatingSystem.__Server;
                    "OSName"             = $OperatingSystem.Caption;
                    "OSVersion"          = $OperatingSystem.version;
                    "MemoryGB"           = $ComputerSystem.TotalPhysicalMemory / 1GB -as [int];
                    "NumberOfProcessors" = $ComputerSystem.NumberOfProcessors;
                    "NumberOfSockets"    = $Sockets;
                    "NumberOfCores"      = $Cores
                }

                $output = New-Object -TypeName PSObject -Property $Info
                $output
            } #end IF Everything_is_OK
        }#end Foreach $Computer in $ComputerName
    }#PROCESS BLOCK
    END {
        # Cleanup
        Write-Verbose -Message "END - Cleanup Variables"
        Remove-Variable -Name output, info, ProcessError, Sockets, Cores, OperatingSystem, ComputerSystem, Processors,
        ComputerName, ComputerName, Computer, Everything_is_OK -ErrorAction SilentlyContinue
        
        # End
        Write-Verbose -Message "END - Script End !"
    }#END BLOCK
}
## End Get-ComputerInfo
## Begin Get-ComputerOS
Function Get-ComputerOS {
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
        [Alias("CN", "__SERVER", "PSComputerName")]
        [String[]]$ComputerName = $env:ComputerName,
		
        [Parameter(ParameterSetName = "Main")]
        [Alias("RunAs")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,
		
        [Parameter(ParameterSetName = "CimSession")]
        [Microsoft.Management.Infrastructure.CimSession]$CimSession
    )
    BEGIN {
        # Default Verbose/Debug message
        Function Get-DefaultMessage {
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
    PROCESS {
        FOREACH ($Computer in $ComputerName) {
            TRY {
                Write-Verbose -Message (Get-DefaultMessage -Message $Computer)
                IF (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                    # Define Hashtable to hold our properties
                    $Splatting = @{
                        class       = "Win32_OperatingSystem"
                        ErrorAction = Stop
                    }
					
                    IF ($PSBoundParameters['CimSession']) {
                        Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - CimSession")
                        # Using cim session already opened
                        $Query = Get-CIMInstance @Splatting -CimSession $CimSession
                    }
                    ELSE {
                        # Credential specified
                        IF ($PSBoundParameters['Credential']) {
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
                        ComputerName    = $Computer
                        OperatingSystem = $Query.Caption
                    }
					
                    # Output
                    New-Object -TypeName PSObject -Property $Properties
                }
            }
            CATCH {
                Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Issue to connect")
                Write-Verbose -Message $Error[0].Exception.Message
            }#CATCH
            FINALLY {
                $Splatting.Clear()
            }
        }#FOREACH
    }#PROCESS
    END {
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
        [parameter(Position = 0, ValueFromPipeline = $true, HelpMessage = "Computer or IP address of machine to test")] 
        [string[]]$ComputerName = $env:COMPUTERNAME, 
        [parameter(HelpMessage = "Pass an alternate credential")] 
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
        foreach ($computer in $computernames) { 
            $WMISplat.ComputerName = $computer 
            try { 
                $wmibios = Get-WmiObject Win32_BIOS @WMISplat -ErrorAction Stop | Select-Object version, serialnumber 
                $wmisystem = Get-WmiObject Win32_ComputerSystem @WMISplat -ErrorAction Stop | Select-Object model, manufacturer
                $ResultProps = @{
                    ComputerName = $computer 
                    BIOSVersion  = $wmibios.Version 
                    SerialNumber = $wmibios.serialnumber 
                    Manufacturer = $wmisystem.manufacturer 
                    Model        = $wmisystem.model 
                    IsVirtual    = $false 
                    VirtualType  = $null 
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
                    if ($wmisystem.manufacturer -like "*Microsoft*") { 
                        $ResultProps.IsVirtual = $true 
                        $ResultProps.VirtualType = "Virtual - Hyper-V" 
                    } 
                    elseif ($wmisystem.manufacturer -like "*VMWare*") { 
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
## Begin Get-DefragAnalysis
Function Get-DefragAnalysis {
    <# .Synopsis
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

    [cmdletbinding(SupportsShouldProcess = $True)]

    Param(
        [Parameter(Position = 0, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullorEmpty()]
        [Alias("drive")]
        [string]$Driveletter = "C:",
        [Parameter(Position = 1, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullorEmpty()]
        [Alias("PSComputername", "SystemName")]
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin {
        Write-Verbose -Message "$(Get-Date) Starting $($MyInvocation.Mycommand)"   
    } #close Begin

    Process {
        #strip off any extra spaces on the drive letter just in case
        Write-Verbose "$(Get-Date) Processing $Driveletter"
        $Driveletter = $Driveletter.Trim()
        if ($Driveletter.length -gt 2) {
            Write-Verbose "$(Get-Date) Scrubbing drive parameter value"
            $Driveletter = $Driveletter.Substring(0, 2)
        }
        #add a colon if not included
        if ($Driveletter -match "^\w$") {
            Write-Verbose "$(Get-Date) Modifying drive parameter value"
            $Driveletter = "$($Driveletter):"
        }

        Write-Verbose "$(Get-Date) Analyzing drive $Driveletter"
        
        Foreach ($computer in $computername) {
            Write-Verbose "$(Get-Date) Examining $computer"
            Try {
                $volume = Get-WmiObject -Class Win32_Volume -filter "DriveLetter='$Driveletter'" -computername $computer -errorAction "Stop"
            }
            Catch {
                Write-Warning ("Failed to get volume {0} from  {1}. {2}" -f $driveletter, $computer, $_.Exception.Message)
            }
            if ($volume) {
                Write-Verbose "$(Get-Date) Running defrag analysis"
                $analysis = $volume | Invoke-WMIMethod -name DefragAnalysis
        
                #get properties for DefragAnalysis so we can filter out system properties
                $analysis.DefragAnalysis.Properties | 
                ForEach-Object { $_.Name }
        
                Write-Verbose "$(Get-Date) Retrieving results"
                $analysis | Select-Object @{Name = "Results"; Expression = { $_.DefragAnalysis | 
                        Select-Object -Property $Prop |
                        Foreach-Object { 
                            #Add on some additional property values
                            $_ | Add-member -MemberType Noteproperty -Name Driveletter -value $DriveLetter
                            $_ | Add-member -MemberType Noteproperty -Name DefragRecommended -value $analysis.DefragRecommended 
                            $_ | Add-member -MemberType Noteproperty -Name Computername -value $volume.__SERVER -passthru
                        } #foreach-object
                    }
                }  | Select-Object -expand Results 
            
                #clean up variables so there are no accidental leftovers
                Remove-Variable "volume", "analysis"
            } #close if volume
        } #close Foreach computer
    } #close Process
 
    End {
        Write-Verbose "$(Get-Date) Defrag analysis complete"
    } #close End

}
## End Get-DefragAnalysis
## Begin Get-DirectoryVolume
Function Get-DirectoryVolume {

    [CmdletBinding()]
    param
    (
        [parameter(
            position = 0,
            mandatory = 1,
            valuefrompipeline = 1,
            valuefrompipelinebypropertyname = 1)]
        [string[]]
        $Path,

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

    process {
        $path `
        | ForEach-Object {
            Write-Verbose ("Checking path : {0}. Scale : {1}. Recurse switch : {2}. Decending : {3}" -f $_, $Scale, $Recurse, !$Ascending)
            if (Test-Path $_) {
                $result = Get-ChildItem -Path $_ -Recurse:$Recurse `
                | Where-Object PSIsContainer `
                | ForEach-Object {
                    $subFolderItems = (Get-ChildItem $_.FullName | Where-Object Length | Measure-Object Length -sum)
                    [PSCustomObject]@{
                        Fullname = $_.FullName
                        $scale   = [decimal]("{0:N4}" -f ($subFolderItems.sum / "1{0}" -f $scale))
                    } } `
                | Sort-Object $scale -Descending:(!$Ascending)

                if ($OmitZero) {
                    return $result | Where-Object $Scale -ne ([decimal]("{0:N4}" -f "0.0000"))
                }
                else {
                    return $result
                }
            }
        }
    }
}
## End Get-DirectoryVolume
## Begin Get-GitCurrentRelease
Function Get-GitCurrentRelease {
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
## End Get-GitCurrentRelease
$volume | sort MB -Descending | ft -AutoSize
18:53

#>

    [cmdletbinding()]
    Param(
        [ValidateNotNullorEmpty()]
        [string]$Uri = "https://api.github.com/repos/git-for-windows/git/releases/latest"
    )
 
    Begin {
        Write-Verbose "[BEGIN  ] Starting: $($MyInvocation.Mycommand)"  
 
    } #begin
    ## End Get-GitCurrentRelease
 
    Process {
        Write-Verbose "[PROCESS] Getting current release information from $uri"
        $data = Invoke-Restmethod -uri $uri -Method Get
 
    
        if ($data.tag_name) {
            [pscustomobject]@{
                Name     = $data.name
                Version  = $data.tag_name
                Released = $($data.published_at -as [datetime])
            }
        } 
    } #process
    ## End Get-GitCurrentRelease
 
    End {
        Write-Verbose "[END    ] Ending: $($MyInvocation.Mycommand)"
    } #end
    ## End Get-GitCurrentRelease
 
}
## End Get-GitCurrentRelease
## Begin Get-HashTableEmptyValue
Function Get-HashTableEmptyValue {
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
        if ($HashTable[$_] -eq "" -or $HashTable[$_] -eq $null) {
            Write-Output $_
        }
    }
}
## End Get-HashTableEmptyValue
## Begin Get-HashTableNotEmptyOrNullValue
Function Get-HashTableNotEmptyOrNullValue {
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
        if ($HashTable[$_] -ne "") {
            Write-Output $_
        }
    }
}
## End Get-HashTableNotEmptyOrNullValue
## Begin Get-HashTableNotEmptyOrNullValue
## Begin Get-ImageInformation
Function Get-ImageInformation {
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
    Foreach ($Image in $FilePath) {
        # Load Assembly
        Add-type -AssemblyName System.Drawing
		
        # Retrieve information
        New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Image
    }
}
## End Get-ImageInformation
## Begin Get-InstalledSoftware
Function Get-InstalledSoftware {
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
    Param
    (
        [Alias('Computer', 'ComputerName', 'HostName')]
        [Parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $true, Position = 1)]
        [string[]]$Name = $env:COMPUTERNAME
    )
    Begin {
        $LMkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall", "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        $LMtype = [Microsoft.Win32.RegistryHive]::LocalMachine
        $CUkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
        $CUtype = [Microsoft.Win32.RegistryHive]::CurrentUser
		
    }
    Process {
        ForEach ($Computer in $Name) {
            $MasterKeys = @()
            If (!(Test-Connection -ComputerName $Computer -count 1 -quiet)) {
                Write-Error -Message "Unable to contact $Computer. Please verify its network connectivity and try again." -Category ObjectNotFound -TargetObject $Computer
                Break
            }
            $CURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($CUtype, $computer)
            $LMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($LMtype, $computer)
            ForEach ($Key in $LMkeys) {
                $RegKey = $LMRegKey.OpenSubkey($key)
                If ($RegKey -ne $null) {
                    ForEach ($subName in $RegKey.getsubkeynames()) {
                        foreach ($sub in $RegKey.opensubkey($subName)) {
                            $MasterKeys += (New-Object PSObject -Property @{
                                    "ComputerName"     = $Computer
                                    "Name"             = $sub.getvalue("displayname")
                                    "SystemComponent"  = $sub.getvalue("systemcomponent")
                                    "ParentKeyName"    = $sub.getvalue("parentkeyname")
                                    "Version"          = $sub.getvalue("DisplayVersion")
                                    "UninstallCommand" = $sub.getvalue("UninstallString")
                                })
                        }
                    }
                }
            }
            ForEach ($Key in $CUKeys) {
                $RegKey = $CURegKey.OpenSubkey($Key)
                If ($RegKey -ne $null) {
                    ForEach ($subName in $RegKey.getsubkeynames()) {
                        foreach ($sub in $RegKey.opensubkey($subName)) {
                            $MasterKeys += (New-Object PSObject -Property @{
                                    "ComputerName"     = $Computer
                                    "Name"             = $sub.getvalue("displayname")
                                    "SystemComponent"  = $sub.getvalue("systemcomponent")
                                    "ParentKeyName"    = $sub.getvalue("parentkeyname")
                                    "Version"          = $sub.getvalue("DisplayVersion")
                                    "UninstallCommand" = $sub.getvalue("UninstallString")
                                })
                        }
                    }
                }
            }
            $MasterKeys = ($MasterKeys | Where-Object { $null -ne $_.Name -AND $_.SystemComponent -ne "1" -AND $null -eq $_.ParentKeyName -eq $Null } | Select-Object Name, Version, ComputerName, UninstallCommand | Sort-Object Name)
            $MasterKeys
        }
    }
    End {
		
    }
}
## End Get-InstalledSoftware
## Begin Get-ISEShortCut
Function Get-ISEShortCut {
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
    PARAM($Key, $Name)
    BEGIN {
        Function Test-IsISE {
            # try...catch accounts for:
            # Set-StrictMode -Version latest
            try {
                return $null -ne $psISE;
            }
            catch {
                return $false;
            }
        }
    }
    PROCESS {
        if ($(Test-IsISE) -eq $true) {
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
## End Get-ISEShortCut

## Begin Get-LogFast
Function Get-LogFast {
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
    BEGIN {
        # Create a StreamReader object
        #  Fortunately this .NET Framework called System.IO.StreamReader allows you to read text files a line at a time which is important when you� re dealing with huge log files :-)
        $StreamReader = New-object -TypeName System.IO.StreamReader -ArgumentList (Resolve-Path -Path $Path -ErrorAction Stop).Path
    }
    PROCESS {
        # .Peek() Method: An integer representing the next character to be read, or -1 if no more characters are available or the stream does not support seeking.
        while ($StreamReader.Peek() -gt -1) {
            # Read the next line
            #  .ReadLine() method: Reads a line of characters from the current stream and returns the data as a string.
            $Line = $StreamReader.ReadLine()
			
            #  Ignore empty line and line starting with a #
            if ($Line.length -eq 0 -or $Line -match "^#") {
                continue
            }
			
            IF ($PSBoundParameters['Match']) {
                If ($Line -match $Match) {
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
## End Get-LogFast

## Begin Get-MachineType
Function Get-MachineType {
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
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param
    (
        # ComputerName
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin {
    }
    Process {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Checking $Computer"
            try {
                # Check to see if $Computer resolves DNS lookup successfuly.
                $null = [System.Net.DNS]::GetHostEntry($Computer)
                
                $ComputerSystemInfo = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop -Credential $Credential
                
                switch ($ComputerSystemInfo.Model) {
                    
                    # Check for Hyper-V Machine Type
                    "Virtual Machine" {
                        $MachineType = "VM"
                    }

                    # Check for VMware Machine Type
                    "VMware Virtual Platform" {
                        $MachineType = "VM"
                    }

                    # Check for Oracle VM Machine Type
                    "VirtualBox" {
                        $MachineType = "VM"
                    }

                    # Check for Xen
                    "HVM domU" {
                        $MachineType = "VM"
                    }

                    # Check for KVM
                    # I need the values for the Model for which to check.

                    # Otherwise it is a physical Box
                    default {
                        $MachineType = "Physical"
                    }
                }
                
                # Building MachineTypeInfo Object
                $MachineTypeInfo = New-Object -TypeName PSObject -Property ([ordered]@{
                        ComputerName = $ComputerSystemInfo.PSComputername
                        Type         = $MachineType
                        Manufacturer = $ComputerSystemInfo.Manufacturer
                        Model        = $ComputerSystemInfo.Model
                    })
                $MachineTypeInfo
            }
            catch [Exception] {
                Write-Output "$Computer`: $($_.Exception.Message)"
            }
        }
    }
    End {

    }
}
## End Get-MachineType
## Begin Get-MappedDrives
Function Get-MappedDrives($ComputerName) {
    #Ping remote machine, continue if available
    $ComputerName = Get-Content -Path C:\LazyWinAdmin\Logs\Contoso.corp\Computers\Windows7.txt
    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
        #Get remote explorer session to identify current user
        $explorer = Get-WmiObject -ComputerName $ComputerName -Class win32_process | Where-Object { $_.name -eq "explorer.exe" }
    
        #If a session was returned check HKEY_USERS for Network drives under their SID
        if ($explorer) {
            $Hive = 2147483651
            $sid = ($explorer.GetOwnerSid()).sid
            $owner = $explorer.GetOwner()
            $RegProv = get-WmiObject -List -Namespace "root\default" -ComputerName $ComputerName | Where-Object { $_.Name -eq "StdRegProv" }
            $DriveList = $RegProv.EnumKey($Hive, "$($sid)\Network")
      
            #If the SID network has mapped drives iterate and report on said drives
            if ($DriveList.sNames.count -gt 0) {
                "$($owner.Domain)\$($owner.user) on $($ComputerName)"
                foreach ($drive in $DriveList.sNames) {
                    "$($drive)`t$(($RegProv.GetStringValue($Hive, "$($sid)\Network\$($drive)", "RemotePath")).sValue)"
                }
            }
            else { "No mapped drives on $($ComputerName)" }
        }
        else { "explorer.exe not running on $($ComputerName)" }
    }
    else { "Can't connect to $($ComputerName)" }

    Out-File -FilePath "C:\LazyWinAdmin\Logs\Contoso.corp\Computers\Mapped\" + $owner + ".txt"
}
## End Get-MappedDrives
## Begin Get-RebootBoolean
Function Get-RebootBoolean {
    Param
    (
        $ComputerName
    )
    Process {
        
        $os = Get-WmiObject win32_operatingsystem -ComputerName $ComputerName
        $uptime = (Get-Date) - ($os.ConvertToDateTime($os.lastbootuptime))
        $minutesUp = $uptime.TotalMinutes
   
    }
    End {
        if ($minutesUp -le 120) {
            return $true
        }
        else {
            return $false
        }
         
    }
}
## End Get-RebootBoolean
## Begin Get-ProcessorBoolean
Function Get-ProcessorBoolean {
    Param
    (
        $ComputerName
    )

    Begin {
    }
    Process {
        $value = (Get-Counter -ComputerName $ComputerName -Counter �\Processor(_Total)\% Processor Time� -SampleInterval 10).CounterSamples.CookedValue
    }
    End {
        if ($value -ge 90) {
            return $true
        }
        else {
            return $false
        }
    }
}
## End Get-ProcessorBoolean
## Begin Get-MemoryBoolean
Function Get-MemoryBoolean {
    Param
    (
        $ComputerName
    )

    Process {
        $value = gwmi -Class win32_operatingsystem -computername $ComputerName | Select-Object @{Name = "MemoryUsage"; Expression = { � { 0:N2 }� -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) * 100) / $_.TotalVisibleMemorySize) } }
    }
    End {
        if ($value.MemoryUsage -ge 90) {
            return $true
        }
        else {
            return $false
        }
        
    }
}
## End Get-MemoryBoolean
## Begin Get-DiskSpaceBoolean
Function Get-DiskSpaceBoolean {
    Param
    (
        $freeBoolean = $false,
        $ComputerName
    )

    Process {
        $diskInfo = Get-WmiObject -ComputerName $ComputerName -class win32_logicaldisk
        foreach ($disk in $diskInfo) {
            if ($disk.DeviceID -ne 'A:') {
                if (($disk.FreeSpace / $disk.Size) * 100 -le 10) {
                    $freeBoolean = $true
                }
            }

        }
    }
    End {
        $freeBoolean
    }
}
## End Get-DiskSpaceBoolean
## Begin Get-NotRunningServices
Function Get-NotRunningServices {
    
    Param
    (
        $ComputerName
    )

    
    Process {
        $notRunning = Get-wmiobject -ComputerName $ComputerName win32_service -Filter "startmode = 'auto' AND state != 'running' AND Exitcode !=0"
        $count = $notRunning.Count
    }
    End {
        if ($count -ge 0) {
            return $true
        }
        else {
            return $false
        }
    }
}
## End Get-NotRunningServices
## Begin Get-Accelerators
Function Get-Accelerators {
    [psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::get
}
## End Get-Accelerators
## Begin Get-NetFramework
Function Get-NetFramework {
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
        [System.Management.Automation.PSCredential]$Credential
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
            ComputerName      = "$($env:Computername)$($env:USERDNSDOMAIN)"
            PowerShellVersion = $psversiontable.PSVersion.Major
            NetFramework      = $netFramework
        }
        New-Object -TypeName PSObject -Property $Properties
    }
}
## End Get-NetFramework
## Begin Get-NetFrameworkTypeAccelerator
Function Get-NetFrameworkTypeAccelerator {
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
## End Get-NetFrameworkTypeAccelerator
## Begin Get-NTSystemInfo
Function Get-NTSystemInfo {
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

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $false,
            ValueFromPipeline = $True,
            HelpMessage = "Enter Computer name or IP address to query")]
        [String[]] $ComputerName = 'localhost',
    
        [Parameter(ValueFromPipeline = $True)]
        [System.Management.Automation.PSCredential]$Credential

    )
    BEGIN {}

    PROCESS {

        if ($Credential) {
            foreach ($Computer in $ComputerName) {
                $OSInfo = Get-WmiObject -class Win32_OperatingSystem -ComputerName $Computer -Credential $Credential
                $NetHardwareInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer -Credential $Credential
                $HardwareInfo = Get-WmiObject -Class Win32_Processor -ComputerName $Computer -Credential $Credential
                $NetHardwareInfo1 = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $Computer -Credential $Credential
                $UserSystemInfo = Get-WmiObject  -class win32_computerSystem -ComputerName $Computer -Credential $Credential

                $Props = @{'Hostname'        = $OSInfo.CSName;
                    'OS Version'             = $OSInfo.name;
                    'Build Number'           = $OSInfo.BuildNumber;
                    'Service Pack'           = $OSInfo.ServicePackMajorVersion;
                    'IP Addresses'           = $NetHardwareInfo  | Sort-Object -Property IPAddress | Where-Object -Property IPAddress -GT $NULL | Select-Object -ExpandProperty IPAddress;
                    'Mac Addresses'          = $NetHardwareInfo | Sort-Object -Property MACAddress | Where-Object -Property MACAddress -GT $NULL | Select-Object -ExpandProperty MACAddress;
                    'Network Speed'          = $NetHardwareInfo1 | Where-Object -Property Speed -GT $NULL | Select-Object -ExpandProperty Speed ;
                    'OS Architecture'        = $OSInfo.OSArchitecture;
                    'Processor Architecture' = $HardwareInfo.DataWidth;
                    'Logged-in User'         = $UserSystemInfo.Username;
                }
                $Object = New-Object -TypeName PSObject -Property $Props
                Write-Output $Object | Select-Object -Property Hostname, 'Logged-in User', 'IP Addresses', 'Network Speed', 'Mac Addresses', 'OS Version', 'Build Number', 'Service Pack', 'OS Architecture', 'Processor Architecture'


            }
        }
        else {
            ## End Get-NTSystemInfo
            foreach ($Computer in $ComputerName) {
                $OSInfo = Get-WmiObject -class Win32_OperatingSystem -ComputerName $Computer
                $NetHardwareInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer
                $HardwareInfo = Get-WmiObject -Class Win32_Processor -ComputerName $Computer
                $NetHardwareInfo1 = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $Computer
                $UserSystemInfo = Get-WmiObject  -class win32_computerSystem -ComputerName $Computer
        
                $Props = @{'Hostname'        = $OSInfo.CSName;
                    'OS Version'             = $OSInfo.name;
                    'Build Number'           = $OSInfo.BuildNumber;
                    'Service Pack'           = $OSInfo.ServicePackMajorVersion;
                    'IP Addresses'           = $NetHardwareInfo  | Sort-Object -Property IPAddress | Where-Object -Property IPAddress -GT $NULL | Select-Object -ExpandProperty IPAddress;
                    'Mac Addresses'          = $NetHardwareInfo | Sort-Object -Property MACAddress | Where-Object -Property MACAddress -GT $NULL | Select-Object -ExpandProperty MACAddress;
                    'Network Speed'          = $NetHardwareInfo1 | Where-Object -Property Speed -GT $NULL | Select-Object -ExpandProperty Speed ;
                    'OS Architecture'        = $OSInfo.OSArchitecture;
                    'Processor Architecture' = $HardwareInfo.DataWidth;
                    'Logged-in User'         = $UserSystemInfo.Username;
                }
            }
            
            $Object = New-Object -TypeName PSObject -Property $Props
            Write-Output $Object | Select-Object -Property Hostname, 'Logged-in User', 'IP Addresses', 'Network Speed', 'Mac Addresses', 'OS Version', 'Build Number', 'Service Pack', 'OS Architecture', 'Processor Architecture'

        }
        ## End Get-NTSystemInfo
    }
    ## End Get-NTSystemInfo
}
## End Get-NTSystemInfo
## Begin Get-PendingReboot
Function Get-PendingReboot {
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
    Process {
        Foreach ($Computer in $ComputerName) {
            Try {
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
                If ([Int32]$WMI_OS.BuildNumber -ge 6001) {
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
                If ($ActCompNm -ne $CompNm) {
                    $CompPendRen = $true
                }
				
                ## If PendingFileRenameOperations has a value set $RegValuePFRO variable to $true 
                If ($RegValuePFRO) {
                    $PendFileRename = $true
                }
				
                ## Determine SCCM 2012 Client Reboot Pending Status 
                ## To avoid nested 'if' statements and unneeded WMI calls to determine if the CCM_ClientUtilities class exist, setting EA = 0 
                $CCMClientSDK = $null
                $CCMSplat = @{
                    NameSpace    = 'ROOT\ccm\ClientSDK'
                    Class        = 'CCM_ClientUtilities'
                    Name         = 'DetermineIfRebootPending'
                    ComputerName = $Computer
                    ErrorAction  = 'Stop'
                }
                ## Try CCMClientSDK 
                Try {
                    $CCMClientSDK = Invoke-WmiMethod @CCMSplat
                }
                Catch [System.UnauthorizedAccessException] {
                    $CcmStatus = Get-Service -Name CcmExec -ComputerName $Computer -ErrorAction SilentlyContinue
                    If ($CcmStatus.Status -ne 'Running') {
                        Write-Warning "$Computer`: Error - CcmExec service is not running."
                        $CCMClientSDK = $null
                    }
                }
                Catch {
                    $CCMClientSDK = $null
                }
				
                If ($CCMClientSDK) {
                    If ($CCMClientSDK.ReturnValue -ne 0) {
                        Write-Warning "Error: DetermineIfRebootPending returned error code $($CCMClientSDK.ReturnValue)"
                    }
                    If ($CCMClientSDK.IsHardRebootPending -or $CCMClientSDK.RebootPending) {
                        $SCCM = $true
                    }
                }
				
                Else {
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
                    Computer           = $WMI_OS.CSName
                    CBServicing        = $CBSRebootPend
                    WindowsUpdate      = $WUAURebootReq
                    CCMClientSDK       = $SCCM
                    PendComputerRename = $CompPendRen
                    PendFileRename     = $PendFileRename
                    PendFileRenVal     = $RegValuePFRO
                    RebootPending      = ($CompPendRen -or $CBSRebootPend -or $WUAURebootReq -or $SCCM -or $PendFileRename)
                } | Select-Object @SelectSplat
				
            }
            Catch {
                Write-Warning "$Computer`: $_"
                ## If $ErrorLog, log the file to a user specified location/path 
                If ($ErrorLog) {
                    Out-File -InputObject "$Computer`,$_" -FilePath $ErrorLog -Append
                }
            }
        } ## End Foreach ($Computer in $ComputerName)       
    } ## End Process 
	
    End { } ## End End 
	
}
## End Get-PendingReboot
## Begin Get-PSCredential
Function Get-PSCredential {

    [CmdletBinding()]
    param(
        [parameter(
            position = 0,
            mandatory = 0)]
        [System.Security.SecureString]$credentialpath = (ConvertTo-SecureString "C:\Deployment\Bin\credential.json" -AsPlainText -Force)
    )

    $credential = Get-Content $credentialpath -Raw | ConvertFrom-Json
    $secpasswd = ConvertTo-SecureString $credential.password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($($credential.user), $secpasswd)    

    return $credential
}
## End Get-PSCredential
## Begin New-PSCredential
Function New-PSCredential {

    [CmdletBinding()]
    param(
        [parameter(
            position = 0,
            mandatory = 0)]
        [System.Security.SecureString]$credentialpath = (ConvertTo-SecureString "C:\Deployment\Bin\credential.json" -AsPlainText -Force),

        [parameter(
            position = 1,
            mandatory)]
        $user,

        [parameter(
            position = 2,
            mandatory)]
        [System.Security.SecureString]$password
    )

    [PSCustomObject]@{
        user     = $user
        password = $password
    } | ConvertTo-Json | Out-File -FilePath $credentialpath -Force

}
## End New-PSCredential

## Begin Get-ProductId
Function Get-ProductId {
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $false,
            Position = 1,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$path = @("registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\*")
    )

    begin {
        $list = New-Object 'System.Collections.Generic.List[PSCustomObject]'
    }

    process {
        foreach ($item in $path) {
            # Validate path existence
            if (-not (Test-Path -Path $item)) {
                Write-Error "Path '$item' not found!"
                continue
            }

            $regEntries = Get-ItemProperty -Path $item | Where-Object { $_.DisplayName }
            
            foreach ($entry in $regEntries) {
                $installDate = $null
                if ($entry.InstallDate -match '^\d{8}$') {
                    try {
                        $installDate = [DateTime]::ParseExact($entry.InstallDate, "yyyyMMdd", $null)
                    } catch {
                        Write-Warning "Failed to parse install date for '$($entry.DisplayName)'"
                    }
                }

                $obj = [PSCustomObject]@{
                    DisplayName    = $entry.DisplayName
                    DisplayVersion = $entry.DisplayVersion
                    Publisher      = $entry.Publisher
                    InstallDate    = $installDate
                    ProductId      = $entry.PSChildName -replace "[{}]"
                }

                $list.Add($obj)
            }
        }
    }

    end {
        $list | Sort-Object DisplayName
    }
}
## End Get-ProductId
## Begin Get-PSObjectEmptyOrNullProperty
Function Get-PSObjectEmptyOrNullProperty {
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
    PROCESS {
        $PsObject.psobject.Properties |
        Where-Object { -not $_.value }
    }
}
## End Get-PSObjectEmptyOrNullProperty
## Begin Get-RemoteComputerDisk
Function Get-RemoteComputerDisk {
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
    
    Param
    (
        $RemoteComputerName
    )

    Begin {
        $output = "Drive `t UsedSpace(in GB) `t FreeSpace(in GB) `t TotalSpace(in GB) `n"
    }
    Process {
        $drives = Get-WmiObject Win32_LogicalDisk -ComputerName $RemoteComputerName

        foreach ($drive in $drives) {
            
            $drivename = $drive.DeviceID
            $freespace = [int]($drive.FreeSpace / 1GB)
            $totalspace = [int]($drive.Size / 1GB)
            $usedspace = $totalspace - $freespace
            $output = $output + $drivename + "`t`t" + $usedspace + "`t`t`t`t`t`t" + $freespace + "`t`t`t`t`t`t" + $totalspace + "`n"
        }
    }
    End {
        return $output
    }
}
## End Get-RemoteComputerDisk
## Begin Get-RemoteProgram
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
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0
        )]
        [string[]]
        $ComputerName = $env:COMPUTERNAME,
        [Parameter(Position = 0)]
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
        $SelectProperty = @('ProgramName', 'ComputerName')
        if ($Property) {
            $SelectProperty += $Property
        }
    }

    process {
        foreach ($Computer in $ComputerName) {
            $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer)
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
                }
                else {
                    $Regex = [regex]"(^(.+?\s){3}).*$|(.*)"
                }
                [System.Collections.ArrayList]$Array = @()
            } -Process {
                if ($ExcludeSimilar) {
                    $null = $Array.Add($_)
                }
                else {
                    $_
                }
            } -End {
                if ($ExcludeSimilar) {
                    $Array | Select-Object -Property *, @{
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
## End Get-RemoteProgram
## Begin Get-ScheduledTask
Function Get-ScheduledTask {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$ComputerName = $env:COMPUTERNAME
    )

    Begin {
        $ST = New-Object -ComObject "Schedule.Service"
    }

    Process {
        ForEach ($Computer in $ComputerName) {
            Try {
                # Check if the computer is reachable before connecting
                if (!(Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
                    Write-Warning "Unable to reach $Computer. Skipping..."
                    continue
                }

                $ST.Connect($Computer)
                $Root = $ST.GetFolder("\")

                $Tasks = @($Root.GetTasks(0))
                if ($Tasks.Count -eq 0) {
                    Write-Warning "No scheduled tasks found on $Computer."
                    continue
                }

                $Tasks | ForEach-Object {
                    $xml = ([xml]$_.xml).Task
                    [PSCustomObject]@{
                        ComputerName   = $Computer
                        Task           = $_.Name
                        Author         = $xml.RegistrationInfo.Author
                        RunAs          = $xml.Principals.Principal.UserId
                        Enabled        = $_.Enabled
                        State          = Switch ($_.State) {
                            0 { 'Unknown' }
                            1 { 'Disabled' }
                            2 { 'Queued' }
                            3 { 'Ready' }
                            4 { 'Running' }
                        }
                        LastTaskResult = Switch ($_.LastTaskResult) {
                            0x0        { "Successfully completed" }
                            0x1        { "Incorrect Function called" }
                            0x2        { "File not found" }
                            0xa        { "Environment is not correct" }
                            0x41301    { "Task is currently running" }
                            0x8004130F { "No account information found for the task" }
                            Default    { "Unknown result code: $_" }
                        }
                        Command        = $xml.Actions.Exec.Command
                        Arguments      = $xml.Actions.Exec.Arguments
                        StartDirectory = $xml.Actions.Exec.WorkingDirectory
                        Hidden         = $xml.Settings.Hidden
                    }
                }
            } Catch {
                Write-Warning "Error processing ${$Computer}: ${($_.Exception.Message)}"
            }
        }
    }
}
## End Get-ScheduledTask
## Begin Get-ScreenResolution
Function Get-ScreenResolution {            
    [void] [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")            
    [void] [Reflection.Assembly]::LoadWithPartialName("System.Drawing")    
    Add-Type -AssemblyName System.Windows.Forms        
    $Screens = [system.windows.forms.screen]::AllScreens            

    foreach ($Screen in $Screens) {            
        $DeviceName = $Screen.DeviceName            
        $Width = $Screen.Bounds.Width            
        $Height = $Screen.Bounds.Height            
        $IsPrimary = $Screen.Primary            

        $OutputObj = New-Object -TypeName PSobject             
        $OutputObj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $DeviceName            
        $OutputObj | Add-Member -MemberType NoteProperty -Name Width -Value $Width            
        $OutputObj | Add-Member -MemberType NoteProperty -Name Height -Value $Height            
        $OutputObj | Add-Member -MemberType NoteProperty -Name IsPrimaryMonitor -Value $IsPrimary            
        $OutputObj            

    }            
    ## End Get-ScreenResolution
}
## End Get-ScreenResolution
## Begin Get-ScreenShot
Function Get-ScreenShot {
    [CmdletBinding()]
    param(
        [parameter(Position = 0, Mandatory = 0, ValueFromPipelinebyPropertyName = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$OutPath = "$env:USERPROFILE\Documents\ScreenShot",

        #screenshot_[yyyyMMdd_HHmmss_ffff].png
        [parameter(Position = 1, Mandatory = 0, ValueFromPipelinebyPropertyName = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$FileNamePattern = 'screenshot_{0}.png',

        [parameter(Position = 2, Mandatory = 0, ValueFromPipeline = 1, ValueFromPipelinebyPropertyName = 1)]
        [ValidateNotNullOrEmpty()]
        [int]$RepeatTimes = 0,

        [parameter(Position = 3, Mandatory = 0, ValueFromPipelinebyPropertyName = 1)]
        [ValidateNotNullOrEmpty()]
        [int]$DurationMs = 1
    )

    begin {
        $ErrorActionPreference = 'Stop'
        Add-Type -AssemblyName System.Windows.Forms

        if (-not (Test-Path $OutPath)) {
            New-Item $OutPath -ItemType Directory -Force
        }
    }

    process {
        1..$RepeatTimes `
        | ForEach-Object {
            $fileName = $FileNamePattern -f (Get-Date).ToString('yyyyMMdd_HHmmss_ffff')
            $path = Join-Path $OutPath $fileName

            $b = New-Object System.Drawing.Bitmap([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width, [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height)
            $g = [System.Drawing.Graphics]::FromImage($b)
            $g.CopyFromScreen((New-Object System.Drawing.Point(0, 0)), (New-Object System.Drawing.Point(0, 0)), $b.Size)
            $g.Dispose()
            $b.Save($path)

            if ($RepeatTimes -ne 0) {
                Start-Sleep -Milliseconds $DurationMs
            }
        }
    }
}
## End Get-ScreenShot
## Begin Get-ScriptAlias
Function Get-ScriptAlias {
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
    PROCESS {
        FOREACH ($File in $Path) {
            TRY {
                # Retrieve file content
                $ScriptContent = (Get-Content $File -Delimiter $([char]0))
				
                # AST Parsing
                $AbstractSyntaxTree = [System.Management.Automation.Language.Parser]::
                ParseInput($ScriptContent, [ref]$null, [ref]$null)
				
                # Find Aliases
                $AbstractSyntaxTree.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true) |
                ForEach-Object -Process {
                    $Command = $_.CommandElements[0]
                    if ($Alias = Get-Alias | Where-Object { $_.Name -eq $Command }) {
						
                        # Output information
                        [PSCustomObject]@{
                            File              = $File
                            Alias             = $Alias.Name
                            Definition        = $Alias.Definition
                            StartLineNumber   = $Command.Extent.StartLineNumber
                            EndLineNumber     = $Command.Extent.EndLineNumber
                            StartColumnNumber = $Command.Extent.StartColumnNumber
                            EndColumnNumber   = $Command.Extent.EndColumnNumber
                            StartOffset       = $Command.Extent.StartOffset
                            EndOffset         = $Command.Extent.EndOffset
							
                        }#[PSCustomObject]
                    }#if ($Alias)
                }#ForEach-Object
            }#TRY
            CATCH {
                Write-Error -Message $($Error[0].Exception.Message)
            } #CATCH
        }#FOREACH ($File in $Path)
    } #PROCESS
}
## End Get-ScriptAlias
## Begin Get-ScriptDirectory
Function Get-ScriptDirectory {
    <#
.SYNOPSIS
   This Function retrieve the current folder path
.DESCRIPTION
   This Function retrieve the current folder path
#>
    if ($null -ne $hostinvocation) {
        Split-Path $hostinvocation.MyCommand.path
    }
    else {
        Split-Path $script:MyInvocation.MyCommand.Path
    }
}
## End Get-ScriptDirectory
## Begin Get-SecurityUpdate
Function Get-SecurityUpdate {
    [Cmdletbinding()] 
    Param( 
        [Parameter()] 
        [String[]]$Computername = $Computername
    )              
    ForEach ($Computer in $Computername) { 
        $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
        ForEach ($Path in $Paths) { 
            #Create an instance of the Registry Object and open the HKLM base key 
            Try { 
                $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $Computer) 
            }
            Catch { 
                $_ 
                Continue 
            } 
            Try {
                #Drill down into the Uninstall key using the OpenSubKey Method 
                $regkey = $reg.OpenSubKey($Path)  
                #Retrieve an array of string that contain all the subkey names 
                $subkeys = $regkey.GetSubKeyNames()      
                #Open each Subkey and use GetValue Method to return the required values for each 
                ForEach ($key in $subkeys) {   
                    $thisKey = $Path + "\\" + $key   
                    $thisSubKey = $reg.OpenSubKey($thisKey)   
                    # prevent Objects with empty DisplayName 
                    $DisplayName = $thisSubKey.getValue("DisplayName")
                    If ($DisplayName -AND $DisplayName -match '^Update for|rollup|^Security Update|^Service Pack|^HotFix') {
                        $Date = $thisSubKey.GetValue('InstallDate')
                        If ($Date) {
                            Write-Verbose $Date 
                            $Date = $Date -replace '(\d{4})(\d{2})(\d{2})', '$1-$2-$3'
                            Write-Verbose $Date 
                            $Date = Get-Date $Date
                        } 
                        If ($DisplayName -match '(?<DisplayName>.*)\((?<KB>KB.*?)\).*') {
                            $DisplayName = $Matches.DisplayName
                            $HotFixID = $Matches.KB
                        }
                        Switch -Wildcard ($DisplayName) {
                            "Service Pack*" { $Description = 'Service Pack' }
                            "Hotfix*" { $Description = 'Hotfix' }
                            "Update*" { $Description = 'Update' }
                            "Security Update*" { $Description = 'Security Update' }
                            Default { $Description = 'Unknown' }
                        }
                        # create New Object with empty Properties 
                        $Object = [pscustomobject] @{
                            Type        = $Description
                            HotFixID    = $HotFixID
                            InstalledOn = $Date
                            Description = $DisplayName
                        }
                        $Object
                    } 
                }   
                $reg.Close() 
            }
            Catch {}                  
        }  
    }  
}
## End Get-SecurityUpdate
## Begin Get-Serial
Function Get-Serial { 
    ######################################################################   
    # Powershell script to get the the serial numbers on remote servers   
    # It will give the serial numbers on remote servers and export to csv 
    # Customized script useful to every one   
    # Please contact  mllsatyanarayana@gmail.com for any suggestions#   
    #########################################################################  
    ####################serial start################# 
    param( 
        $computername = $env:computername 
    ) 
 
    $os = Get-WmiObject Win32_bios -ComputerName $computername -ea silentlycontinue 
    if ($os) { 
 
        $SerialNumber = $os.SerialNumber  
        $servername = $os.PSComputerName   
        $results = new-object psobject 
        $results | Add-Member noteproperty SerialNumber  $SerialNumber 
        $results | Add-Member noteproperty ComputerName  $servername 
        #Display the results 
 
        $results | Select-Object computername, SerialNumber 
    }  
    else { 
        $results = New-Object psobject 
        $results = New-object psobject 
        $results | Add-Member noteproperty SerialNumber "Na" 
        $results | Add-Member noteproperty ComputerName $servername 
  
        #display the results 
 
        $results | Select-Object computername, SerialNumber 
  
    }  
 
    $infserial = @()  
    foreach ($allserver in $allservers) { 
        $infserial += Get-Serial $allserver  
    } 
    $infserial  
}
## End Get-Serial
## Begin Get-Servers
Function Get-Servers {
    Param (
        [parameter(Mandatory = $False)]
        [ValidateSet('CORP', 'SVC', 'RES', 'PROD')]
        [string]$Domain = 'CORP'
    )
    $Searcher = [adsisearcher]""
    If ($Domain = "SVC") { $searchroot = [ADSI]"LDAP://DC=svc,DC=prod,DC=vegas,DC=com" }
    elseif ($Domain = "PROD") { $searchroot = [ADSI]"LDAP://DC=prod,DC=vegas,DC=com" }
    elseif ($Domain = "RES") { $searchroot = [ADSI]"LDAP://DC=res,DC=vegas,DC=com" }
    else { $searchroot = [ADSI]"LDAP://DC=corp,DC=vegas,DC=com" }
    $Searcher.SearchRoot = $searchroot
    $Searcher.Filter = '(&(objectCategory=computer)(OperatingSystem=Windows*Server*))'
    $Searcher.pagesize = 10000
    $Searcher.sizelimit = 50000
    $Searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach-Object { $_.Properties.dnshostname }
}
## End Get-Servers
## Begin Get-NetworkLevelAuthentication
Function Get-NetworkLevelAuthentication {
    <#
	.SYNOPSIS
		This Function will get the NLA setting on a local machine or remote machine
		# Blog Article: http://lazywinadmin.com/2014/04/powershell-getset-network-level.html
		# GitHub : https://github.com/lazywinadmin/PowerShell/blob/master/TOOL-Get-Set-NetworkLevelAuthentication/Get-Set-NetworkLevelAuthentication.ps1

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
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential
    )#Param
    BEGIN {
        TRY {
            IF (-not (Get-Module -Name CimCmdlets)) {
                Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
                Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
            }
        }
        CATCH {
            IF ($ErrorBeginCimCmdlets) {
                Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
            }
        }
    }#BEGIN
	
    PROCESS {
        FOREACH ($Computer in $ComputerName) {
            TRY {
                # Building Splatting for CIM Sessions
                $CIMSessionParams = @{
                    ComputerName  = $Computer
                    ErrorAction   = 'Stop'
                    ErrorVariable = 'ProcessError'
                }
				
                # Add Credential if specified when calling the Function
                IF ($PSBoundParameters['Credential']) {
                    $CIMSessionParams.credential = $Credential
                }
				
                # Connectivity Test
                Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
                Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
                # CIM/WMI Connection
                #  WsMAN
                IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0') {
                    Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
                    $CimSession = New-CimSession @CIMSessionParams
                    $CimProtocol = $CimSession.protocol
                    Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
                }
				
                # DCOM
                ELSE {
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
                    'ComputerName'     = $NLAinfo.PSComputerName
                    'NLAEnabled'       = $NLAinfo.UserAuthenticationRequired -as [bool]
                    'TerminalName'     = $NLAinfo.TerminalName
                    'TerminalProtocol' = $NLAinfo.TerminalProtocol
                    'Transport'        = $NLAinfo.transport
                }
            }
			
            CATCH {
                Write-Warning -Message "PROCESS - Error on $Computer"
                $_.Exception.Message
                if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
                if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
            }#CATCH
        } # FOREACH
    }#PROCESS
    END {
		
        if ($CimSession) {
            Write-Verbose -Message "END - Close CIM Session(s)"
            Remove-CimSession $CimSession
        }
        Write-Verbose -Message "END - Script is completed"
    }
}
## End Get-NetworkLevelAuthentication
## Begin Get-Skew
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
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Servers 
       
    )

    $RemoteServer = Get-Time -ServerName $Servers
    $LocalServer = Get-Time -ServerName LASDC01
 
    $Skew = $LocalServer.DateTime - $RemoteServer.DateTime
 
    # Check if the time is over 30 seconds
    If (($Skew.TotalSeconds -gt 30) -or ($Skew.TotalSeconds -lt -30)) {
        Write-Host "Time is not within 30 seconds"
    }
    Else {
        ## End Get-Skew
        Write-Host "Time checked ok"
    }
    ## End Get-Skew
}
## End Get-Skew
## Begin Get-Software
Function Get-Software {
    [OutputType('System.Software.Inventory')]
    [Cmdletbinding()] 
    Param( 
        [Parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)] 
        [String[]]$Computername = $env:COMPUTERNAME
    )         
    Begin {
    }
    Process {     
        ForEach ($Computer in $Computername) { 
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
                ForEach ($Path in $Paths) { 
                    Write-Verbose "Checking Path: $Path"
                    # Create an instance of the Registry Object and open the HKLM base key 
                    Try { 
                        $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $Computer, 'Registry64') 
                    }
                    Catch { 
                        Write-Error $_ 
                        Continue 
                    } 
                    # Drill down into the Uninstall key using the OpenSubKey Method 
                    Try {
                        $regkey = $reg.OpenSubKey($Path)  
                        # Retrieve an array of string that contain all the subkey names 
                        $subkeys = $regkey.GetSubKeyNames()      
                        # Open each Subkey and use GetValue Method to return the required values for each 
                        ForEach ($key in $subkeys) {   
                            Write-Verbose "Key: $Key"
                            $thisKey = $Path + "\\" + $key 
                            Try {  
                                $thisSubKey = $reg.OpenSubKey($thisKey)   
                                # Prevent Objects with empty DisplayName 
                                $DisplayName = $thisSubKey.getValue("DisplayName")
                                If ($DisplayName -AND $DisplayName -notmatch '^Update for|rollup|^Security Update|^Service Pack|^HotFix') {
                                    $Date = $thisSubKey.GetValue('InstallDate')
                                    If ($Date) {
                                        Try {
                                            $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)
                                        }
                                        Catch {
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
                                        $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32, 0)))
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
                                    # Removed unused variable $HelpLink
                                    $Object = [pscustomobject]@{
                                        Computername    = $Computer
                                        DisplayName     = $DisplayName
                                        Version         = $Version
                                        InstallDate     = $Date
                                        Publisher       = $Publisher
                                        UninstallString = $UninstallString
                                        InstallLocation = $InstallLocation
                                        InstallSource   = $InstallSource
                                        HelpLink        = $thisSubKey.GetValue('HelpLink')
                                        EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize') * 1024) / 1MB, 2))
                                    }
                                    $Object.pstypenames.insert(0, 'System.Software.Inventory')
                                    Write-Output $Object
                                }
                            }
                            Catch {
                                Write-Warning "$Key : $_"
                            }   
                        }
                    }
                    Catch {}   
                    $reg.Close() 
                }                  
            }
            Else {
                Write-Error "$($Computer): unable to reach remote system!"
            }
        } 
    } 
} 
## End Get-Software
## Begin Get-StrictMode
Function Get-StrictMode {

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
    while ($currentScope) {
        $strictModeVersion = $currentScope | Get-Field -name *StrictModeVersion* -valueOnly
        $currentScope = $currentScope | Get-Field -name *Parent* -valueOnly

        if ($showAllScopes) {
            New-Object PSObject -Property @{
                Scope             = $scope++
                StrictModeVersion = $strictModeVersion
            }
        }
        elseif ($strictModeVersion) {
            $strictModeVersion
        }
    }
}
## End Get-StrictMode
## Begin Get-Field
Function Get-Field {
    [CmdletBinding()]
    param (
        [Parameter(
            mandatory = 0,
            Position = 0)]
        [string[]]
        $name = "*",

        [Parameter(
            mandatory = 1,
            position = 1,
            ValueFromPipeline = 1)]
        $inputObject,
            
        [switch]
        $valueOnly
    )
 
    process {
        $type = $inputObject.GetType()
        [string[]]$bindingFlags = ("Public", "NonPublic", "Instance")

        $type.GetFields($bindingFlags) `
        | Where-Object {
            foreach ($currentName in $name) {
                if ($_.Name -like $currentName) { 
                    return $true
                }
            } } `
        | ForEach-Object {
            $currentField = $_
            $currentFieldValue = $type.InvokeMember(
                $currentField.Name,
                $bindingFlags + "GetField",
                $null,
                $inputObject,
                $null
            )
                
            if ($valueOnly) {
                $currentFieldValue
            }
            else {
                $returnProperties = @{}
                foreach ($prop in @("Name", "IsPublic", "IsPrivate")) {
                    $ReturnProperties.$prop = $CurrentField.$prop
                }

                $returnProperties.Value = $currentFieldValue
                New-Object PSObject -Property $returnProperties
            }
        } 
    }
}
## End Get-Field
## Begin Get-StringCharCount
Function Get-StringCharCount {
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
## End Get-StringCharCount
## Begin Get-StringLastDigit
Function Get-StringLastDigit {
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
    if ($String -match "^.*\d$") {
        # Output the last digit
        $String.Substring(($String.ToCharArray().count) - 1)
    }
    else { Write-Verbose -Message "The following string does not finish by a digit: $String" }
}
## End Get-StringLastDigit
## Begin Get-Time
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
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ServerName,
 
        $Credential
 
    )
    try {
        If ($Credential) {
            $DT = Get-WmiObject -Class Win32_LocalTime -ComputerName $servername -Credential $Credential
        }
        Else {
            $DT = Get-WmiObject -Class Win32_LocalTime -ComputerName $servername
        }
    }
    catch {
        throw
    }
 
    $Times = New-Object PSObject -Property @{
        ServerName = $DT.__Server
        DateTime   = (Get-Date -Day $DT.Day -Month $DT.Month -Year $DT.Year -Minute $DT.Minute -Hour $DT.Hour -Second $DT.Second)
    }
    $Times
 
}
## End Get-Time
## Begin Get-UAC
Function Get-UAC {
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

    begin {
        $path = "registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        $name = "EnableLUA"
    }

    process {
        $uac = Get-ItemProperty $path
        if ($uac.$name -eq 1) { 
            Write-Verbose ("Registry path '{0}', name '{1}', value '{2}'" -f (Get-Item $path).name, $name, $uac.$name)
            "Enabled" 
        } 
        elseif ($uac.$name -eq 0) { 
            Write-Verbose ("Registry path '{0}', name '{1}', value '{2}'" -f (Get-Item $path).name, $name, $uac.$name)
            "Disabled" 
        }
        else {
            Write-Verbose ("Registry path '{0}', name '{1}', value '{2}'" -f (Get-Item $path).name, $name, $uac.$name)
            "Unknown"
        }
    }
}
## End Get-UAC
## Begin Get-UDVariable
Function Get-UDVariable {
    get-variable | where-object { (@(
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
## End Get-UDVariable
## Begin Get-Uptime
Function Get-Uptime {
    $os = Get-WmiObject win32_operatingsystem
    $uptime = (Get-Date) - ($os.ConvertToDateTime($os.lastbootuptime))
    $Display = "Uptime: " + $Uptime.Days + " days, " + $Uptime.Hours + " hours, " + $Uptime.Minutes + " minutes" 
    Write-Output $Display
}
## End Get-Uptime
## Begin Get-UrlRedirection
Function Get-UrlRedirection {
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

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline)] [Uri] $Url,
        [switch] $Enumerate,
        [int] $MaxRedirections = 50 # Use same default as [System.Net.HttpWebRequest]
    )

    process {
        try {

            if ($Enumerate) {
                # Enumerate the whole redirection chain, from input URL to ultimate target,
                # assuming the max. count of redirects is not exceeded.
                # We must walk the chain of redirections one by one.
                # If we disallow redirections, .GetResponse() fails and we must examine
                # the exception's .Response object to get the redirect target.
                $nextUrl = $Url
                $urls = @( $nextUrl.AbsoluteUri ) # Start with the input Uri
                $ultimateFound = $false
                # Note: We add an extra loop iteration so we can determine whether
                #       the ultimate target URL was reached or not.
                foreach ($i in 1..$($MaxRedirections + 1)) {
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
                    }
                    catch [System.Net.WebException] {
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
                    if ($nextUrlStr -match '^https?:') {
                        # absolute URL
                        $nextUrl = $prevUrl = [Uri] $nextUrlStr
                    }
                    else {
                        # URL without scheme and server component
                        $nextUrl = $prevUrl = [Uri] ($prevUrl.Scheme + '://' + $prevUrl.Authority + $nextUrlStr)
                    }
                    if ($i -le $MaxRedirections) { $urls += $nextUrl.AbsoluteUri }          
                }
                # Output the array of URLs (chain of redirections) as a *single* object.
                Write-Output -NoEnumerate $urls
                if (-not $ultimateFound) { Write-Warning "Enumeration of $Url redirections ended before reaching the ultimate target." }

            }
            else {
                # Resolve just to the ultimate target,
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

        }
        catch {
            Write-Error $_ # Report the exception as a non-terminating error.
        }
    } # process

}
## End Get-UrlRedirection

## Begin Get-UserShareDACL
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
                    Name         = $Share.Name
                    Path         = $MoreShare.Path
                    Type         = $ShareType[[int]$MoreShare.Type]
                    Description  = $MoreShare.Description
                    DACLName     = $DACL.Trustee.Name
                    AccessRight  = $AccessMask[[int]$DACL.AccessMask]
                    AccessType   = $AceType[[int]$DACL.AceType]                    
                }
            }
        }
    }
    #Catch any errors                
    Catch {}                                                    
}
## End Get-UserShareDACL
## Begin GetCulture
Function GetCulture {
    Param([int]$WordValue)
	
    #codes obtained from http://support.microsoft.com/kb/221435
    #http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
    $CatalanArray = 1027
    $ChineseArray = 2052, 3076, 5124, 4100
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

    Switch ($WordValue) {
        { $CatalanArray -contains $_ } { $CultureCode = "ca-" }
        { $ChineseArray -contains $_ } { $CultureCode = "zh-" }
        { $DanishArray -contains $_ } { $CultureCode = "da-" }
        { $DutchArray -contains $_ } { $CultureCode = "nl-" }
        { $EnglishArray -contains $_ } { $CultureCode = "en-" }
        { $FinnishArray -contains $_ } { $CultureCode = "fi-" }
        { $FrenchArray -contains $_ } { $CultureCode = "fr-" }
        { $GermanArray -contains $_ } { $CultureCode = "de-" }
        { $NorwegianArray -contains $_ } { $CultureCode = "nb-" }
        { $PortugueseArray -contains $_ } { $CultureCode = "pt-" }
        { $SpanishArray -contains $_ } { $CultureCode = "es-" }
        { $SwedishArray -contains $_ } { $CultureCode = "sv-" }
        Default { $CultureCode = "en-" }
    }
	
    Return $CultureCode
}
## End GetCulture
## Begin Write-Log
Function Write-Log {
    [CmdletBinding()]
    #[Alias('wl')]
    [OutputType([int])]
    Param(
        # The string to be written to the log.
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 0)] [ValidateNotNullOrEmpty()] [Alias("LogContent")] [string]$Message,
        # The path to the log file.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, Position = 1)] [Alias('LogPath')] [string]$Path = $global:DefaultLog,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, Position = 2)] [ValidateSet("Error", "Warn", "Info", "Load", "Execute")] [string]$Level = "Info",
        [Parameter(Mandatory = $false)] [switch]$NoClobber
    )

    Process {
        
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
## End Write-Log
## Begin CheckExists
Function CheckExists {
    param(
        [Parameter(mandatory = $true, position = 0)]$itemtocheck,
        [Parameter(mandatory = $true, position = 1)]$colection
    )
    BEGIN {
        $item = $null
        $exist = $false
    }
    PROCESS {
        foreach ($item in $colection) {
            if ($item.EventID -eq $itemtocheck) {
                $exist = $true
                break;
            }
        }

    }
    END {
        return $exist
    }

}
## End CheckExists
## Begin CheckCount
Function CheckCount {
    param(
        [Parameter(mandatory = $true, position = 0)]$itemtocheck,
        [Parameter(mandatory = $true, position = 1)]$colection
    )
    BEGIN {
        $item = $null
        $count = 0
    }
    PROCESS {
        foreach ($item in $colection) {
			
            if ($item.EventID -eq $itemtocheck) {
                $count++
            }
        }

    }
    END {
        return $count
    }

}
## End CheckCount
## Begin Get-Times
Function Get-Times {
    param(
        [Parameter(mandatory = $true, position = 0)]$colection,
        [Parameter(mandatory = $true, position = 1)]$EventID

    )
    BEGIN {
        $filterCollection = $colection | Where-Object { $_.EventID -eq $EventID }
    }
    PROCESS {
        $previous = $filterCollection[0].TimeWritten
        $last = $filterCollection[0].TimeWritten
        foreach ($item in $filterCollection) {
            if ($item.TimeWritten -lt $previous) {
                $previous = $item.TimeWritten
            }
            if ($item.TimeWritten -gt $last) {
                $last = $item.TimeWritten
            }

        }

    }
    END {
        $output = New-Object psobject -Property @{
            first = $previous
            last  = $last
        }
        return $output
    }

}
## End Get-Times
## Begin Get-EventSubscriber
Function Get-EventSubscriber {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, position = 0)] [int]$Days,
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, position = 1)] [ValidateSet("System", "Security", "Application")][ValidateNotNullOrEmpty()] [String]$LogName,
        [Parameter(Mandatory = $false, ValueFromPipeline = $True, position = 2)] [String]$computer = ".", #dot for localhost it can be changed to get any computer in the domain (server or client)
        [Parameter(Mandatory = $false, ValueFromPipeline = $True, position = 3)] [Switch]$IncludeInfo
    )
    BEGIN {
        if ($LogName -ne "Security") {
            Write-Log -Level Execute -Message "Getting $LogName Events"
        }
        else {
            Write-Log -Level Execute -Message "Getting $LogName Events. This can take a while"
        }
        #In case log is already there remove it.

    }
    
    PROCESS {
        if ($LogName -ne "Security") {
            if ($IncludeInfo) {
                $Log = Get-EventLog -Computername $computer -LogName "$LogName" -EntryType "Information", "Error", "Warning" -After (Get-Date).Adddays(-$Days) | select *
            }
            else {
                $Log = Get-EventLog -Computername $computer -LogName "$LogName" -EntryType "Error", "Warning" -After (Get-Date).Adddays(-$Days) | select *
            }
        }
        else {
            $Log = Get-EventLog -Computername $computer -LogName "$LogName" -EntryType FailureAudit, SuccessAudit -After (Get-Date).Adddays(-$Days) | select *
        }

        $Count = if ($Log.Count) { $log.Count }else { 1 }
        #
        # if($log.EventId -ne $null){
        #     $Count++;
        # }
        # else{
        #      
        # }

        Write-Log -Level Execute -Message "Attaching new properties to $LogName Events. Total Number of items in Log: $Count"       
        $return = @()
        $Log | foreach { $temp = $_.EventID; $valor = CheckCount -itemtocheck $temp -colection $Log; $Dates = Get-Times -colection $Log -EventID $temp;  
            $_ |  Add-Member -Name "Frequency" -Value $valor -MemberType NoteProperty; 
            $_ |  Add-Member -Name "LastTime"  -Value $Dates.Last -MemberType NoteProperty;
            $_ |  Add-Member -Name "FirstTime" -Value $Dates.first -MemberType NoteProperty;
            $i++; $progress = ($i * 100) / $Count;  
            if ($progress -lt 100) { Write-Progress -Activity "Attaching new properties to $LogName Events" -PercentComplete $progress; }
            else { Write-Progress -Activity "Attaching new properties to $LogName Events" -PercentComplete $progress -Complete }
            if (-not (CheckExists $temp $return)) { $return += $_ } }
        
    }
    END {
        return $return | Sort-Object Frequency -Descending
    }
}
## End Get-EventSubscriber
## Begin ObjectsToHtml5
Function ObjectsToHtml5 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, position = 0)][String]$Computer,
        [Parameter(mandatory = $false, position = 1)]$systemObjs,
        [Parameter(mandatory = $false, position = 2)]$AppObjs,
        [Parameter(mandatory = $false, position = 3)]$SecObjs
    )
    BEGIN {
        write-verbose "Setting Actual Date"
        $fecha = get-date -UFormat "%Y%m%d"
        $dia = get-date -UFormat "%A"
        
        $Fn = "$fecha$Filename"
        
        $HtmlFileName = "$global:ScriptLocation\$Filename.html"
        $title = "Event Logs $fecha/$computer"
    }
    PROCESS {
        $html = '<!DOCTYPE HTML>
<html lang="en-US">
<head>
	<meta charset="UTF-8">
	<title>'
        $html += $title
        $html += "</title>
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
        foreach ($item in $systemObjs) {
            $machine = $item.MachineName
            $index = $item.Index
            $timeGenerated = $item.TimeGenerated
            $entrytipe = $item.EntryType
            $Source = $item.Source
            $instanceid = $item.InstanceID
            $Ft = $item.FirstTime
            $Lt = $item.LastTime
            $Eventid = $item.EventID
            $frequency = $item.Frequency
            $mensaje = $item.Message
            $html += "<tr> <td> $machine</td> <td>$index</td> <td>$timeGenerated</td> <td> $entrytipe </td> <td>$Source</td> <td>$instanceid </td><td>$Ft</td> <td>$Lt</td> <td>$Eventid</td> <td>$frequency </td> <td>$mensaje</td> </tr>"
        }
        ## End ObjectsToHtml5

        $html += "
</table>
<h2> Application Information </h2>
<table>
<tr>
<th>MachineName</th><th>Index</th><th>TimeGenerated</th><th>EntryType</th><th>Source</th><th>InstanceID</th><th>FirstTime</th><th>LastTime</th><th>EventID</th><th>Frequency</th><th>Message</th>
</tr>"

        foreach ($item in $AppObjs) {
            $machine = $item.MachineName
            $index = $item.Index
            $timeGenerated = $item.TimeGenerated
            $entrytipe = $item.EntryType
            $Source = $item.Source
            $instanceid = $item.InstanceID
            $Ft = $item.FirstTime
            $Lt = $item.LastTime
            $Eventid = $item.EventID
            $frequency = $item.Frequency
            $mensaje = $item.Message
            $html += "<tr> <td> $machine</td> <td>$index</td> <td>$timeGenerated</td> <td> $entrytipe </td> <td>$Source</td> <td>$instanceid </td><td>$Ft</td> <td>$Lt</td> <td>$Eventid</td> <td>$frequency </td> <td>$mensaje</td> </tr>"
        }
        ## End ObjectsToHtml5

        $html += "
</table>
<h2> Security Information </h2>"

        if ($SecObjs.Count -gt 0) {

            $html += "
<table>
<tr>
<th>MachineName</th><th>Index</th><th>TimeGenerated</th><th>EntryType</th><th>Source</th><th>InstanceID</th><th>FirstTime</th><th>LastTime</th><th>EventID</th><th>Frequency</th><th>Message</th>
</tr>"

            foreach ($item in $SecObjs) {
                $machine = $item.MachineName
                $index = $item.Index
                $timeGenerated = $item.TimeGenerated
                $entrytipe = $item.EntryType
                $Source = $item.Source
                $instanceid = $item.InstanceID
                $Ft = $item.FirstTime
                $Lt = $item.LastTime
                $Eventid = $item.EventID
                $frequency = $item.Frequency
                $mensaje = $item.Message
                $html += "<tr> <td> $machine</td> <td>$index</td> <td>$timeGenerated</td> <td> $entrytipe </td> <td>$Source</td> <td>$instanceid </td><td>$Ft</td> <td>$Lt</td> <td>$Eventid</td> <td>$frequency </td> <td>$mensaje</td> </tr>"

            }
            ## End ObjectsToHtml5
            $html += "</table>"

        }
        ## End ObjectsToHtml5
        else {

            $html += "<p> Security objects not selected in the query. If you want this information please re-run the script with the option '-AddSecurity' <br> Also remeber that you can also add the information information with the switch '-Addinformation'</p>"
        }
        ## End ObjectsToHtml5


        $html += "
	<footer>
	<a href=""https://www.j0rt3g4.com"" target=""_blank"">
	2017 - J0rt3g4 Consulting Services </a> | - &#9400; All rigths reserved.
	</footer>
</body>
</html>"

    
    }
    END {
        $html | Out-File "$global:ScriptLocation\$fecha-$dia-$computername.html" 
    
    }
}
## End ObjectsToHtml5
#Get warnings and info on each event viewer log.
## Begin GetEventErrors
Function GetEventErrors {
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
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, position = 0)] [String]$ScriptPath = ".",
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, position = 1)] [int]$Days = $global:DefaultNumberOfDays, #Day(s) behind for the checking of the logs (default 3) set in line 50
        [Parameter(Mandatory = $false, ValueFromPipeline = $True, position = 2)] [String]$computer = ".", #dot for localhost it can be changed to get any computer in the domain (server or client)
        [Parameter(Mandatory = $false, ValueFromPipeline = $True, position = 3)] [Switch]$AddInfo = $false, #Add Information Events
        [Parameter(Mandatory = $false, ValueFromPipeline = $True, position = 4)] [Switch]$AddSecu = $false  #add security Log
    )

    BEGIN {
        write-Verbose "Preparing script's Variables"
        #SimpleLogname
        $SystemLogName = "System" #other options are: security, application, forwarded events
        $AppLogname = "Application"
        $SecurityLogName = "Security"    
        #set html header in a variable CSS3
        #$header= "<style type=""text/css"">body,html{height:100%}a,abbr,acronym,address,applet,b,big,blockquote,body,caption,center,cite,code,dd,del,dfn,div,dl,dt,em,fieldset,font,form,html,i,iframe,img,ins,kbd,label,legend,li,object,ol,p,pre,q,s,samp,small,span,strike,strong,sub,sup,table,tbody,td,tfoot,th,thead,tr,tt,u,ul,var{margin:0;padding:0;border:0;outline:0;font-size:100%;vertical-align:baseline;background:0 0}body{line-height:1}ol,ul{list-style:none}blockquote,q{quotes:none}blockquote:after,blockquote:before,q:after,q:before{content:'';content:none}:focus{outline:0}del{text-decoration:line-through}table{border-spacing:0;margin: 50px 0px 50px 10px;}body{font-family:Arial,Helvetica,sans-serif;margin:0 15px;width:520px}a:link,a:visited{color:#666;font-weight:700;text-decoration:none}a:active,a:hover{color:#bd5a35;text-decoration:underline}table a:link{color:#666;font-weight:700;text-decoration:none}table a:visited{color:#999;font-weight:700;text-decoration:none}table a:active,table a:hover{color:#bd5a35;text-decoration:underline}table{font-family:Arial,Helvetica,sans-serif;color:#666;font-size:12px;text-shadow:1px 1px 0 #fff;background:#eaebec;border:1px solid #ccc;-moz-border-radius:3px;-webkit-border-radius:3px;border-radius:3px;-moz-box-shadow:0 1px 2px #d1d1d1;-webkit-box-shadow:0 1px 2px #d1d1d1;box-shadow:0 1px 2px #d1d1d1}table th{padding:21px 25px 22px;border-top:1px solid #fafafa;border-bottom:1px solid #e0e0e0;background:#ededed;background:-webkit-gradient(linear,left top,left bottom,from(#ededed),to(#ebebeb));background:-moz-linear-gradient(top,#ededed,#ebebeb)}table th:first-child{text-align:left;padding-left:20px}table tr:first-child th:first-child{-moz-border-radius-topleft:3px;-webkit-border-top-left-radius:3px;border-top-left-radius:3px}table tr:first-child th:last-child{-moz-border-radius-topright:3px;-webkit-border-top-right-radius:3px;border-top-right-radius:3px}table tr{text-align:center;padding-left:20px}table tr td:first-child{text-align:left;padding-left:20px;border-left:0}table tr td{padding:18px;border-top:1px solid #fff;border-bottom:1px solid #e0e0e0;border-left:1px solid #e0e0e0;background:#fafafa;background:-webkit-gradient(linear,left top,left bottom,from(#fbfbfb),to(#fafafa));background:-moz-linear-gradient(top,#fbfbfb,#fafafa)}table tr.even td{background:#f6f6f6;background:-webkit-gradient(linear,left top,left bottom,from(#f8f8f8),to(#f6f6f6));background:-moz-linear-gradient(top,#f8f8f8,#f6f6f6)}table tr:last-child td{border-bottom:0}table tr:last-child td:first-child{-moz-border-radius-bottomleft:3px;-webkit-border-bottom-left-radius:3px;border-bottom-left-radius:3px}table tr:last-child td:last-child{-moz-border-radius-bottomright:3px;-webkit-border-bottom-right-radius:3px;border-bottom-right-radius:3px}table tr:hover td{background:#f2f2f2;background:-webkit-gradient(linear,left top,left bottom,from(#f2f2f2),to(#f0f0f0));background:-moz-linear-gradient(top,#f2f2f2,#f0f0f0);div{font-size:20px;}}</style>";
        #$header= "<style type=""text/css"">{margin:0;padding:0}body{font:14px/1.4 Georgia,Serif}#page-wrap{margin:50px}p{margin:20px 0}table{width:100%;border-collapse:collapse}tr:nth-of-type(odd){background:#eee}th{background:#333;color:#fff;font-weight:700}td,th{padding:6px;border:1px solid #ccc;text-align:left}</style>";
        $header = "<style type=""text/css"">{margin:0;padding:0}@import url(https://fonts.googleapis.com/css?family=Indie+Flower|Josefin+Sans|Orbitron:500|Yrsa);body{font-family:14px/1.4 'Indie Flower','Josefin Sans',Orbitron,sans-serif;font-family:'Indie Flower',cursive;font-family:'Josefin Sans',sans-serif;font-family:Orbitron,sans-serif}#page-wrap{margin:50px}tr:nth-of-type(odd){background:#eee}th{background:#EF5525;color:#fff;font-family:Orbitron,sans-serif}td,th{padding:6px;border:1px solid #ccc;text-align:center;font-size:large}table{width:90%;border-collapse:collapse;margin-left:auto;margin-right:auto;font-family:Yrsa,serif}</style>";
    }
    PROCESS {
        #GET ALL ITEMS in Event Viewer with the selected options
        if (-not $AddSecu -and -not $AddInfo) {
            Write-Log -Level Load -Message "Querying Logs in System and Application Logs without informational items (just Warnings and errors)"
            $system = Get-EventSubscriber -Days $Days -LogName "$SystemLogName" -computer $computerName 
            $appl = Get-EventSubscriber -Days $Days -LogName "$AppLogname" -computer $computerName
        }
        elseif ($AddSecu -and -not $AddInfo) {
            Write-Log -Level Load -Message "Querying Logs in System, Application and security Logs without informational items (just Warnings and errors)"
            $system = Get-EventSubscriber -Days $Days -LogName "$SystemLogName" -computer $computerName 
            $appl = Get-EventSubscriber -Days $Days -LogName "$AppLogname" -computer $computerName
            $security = Get-EventSubscriber -Days $Days -LogName "$SecurityLogName" -computer $computerName
        }
        elseif (-not $AddSecu -and $AddInfo) {
            Write-Log -Level Load -Message "Querying Logs in System and Application Logs WITH informational items"
            $system = Get-EventSubscriber -Days $Days -LogName "$SystemLogName" -computer $computerName -IncludeInfo
            $appl = Get-EventSubscriber -Days $Days -LogName "$AppLogname" -computer $computerName -IncludeInfo
        }
        else {
            Write-Log -Level Load -Message "Querying Logs in System, Application and security Logs WITH informational items (just in system and application, security doesn't have informational items)"
            $system = Get-EventSubscriber -Days $Days -LogName "$SystemLogName" -computer $computerName -IncludeInfo
            $appl = Get-EventSubscriber -Days $Days -LogName "$AppLogname" -computer $computerName -IncludeInfo
            $security = Get-EventSubscriber -Days $Days -LogName "$SecurityLogName" -computer $computerName
        }

        if ($AddSecu) {
            ObjectsToHtml5 -Computer $computerName -systemObjs $system -AppObjs $appl -SecObj $security
        }
        else {
            ObjectsToHtml5 -Computer $computerName -systemObjs $system -AppObjs $appl
        }
    }
    END {
        write-verbose "Done Exporting"
    }
}
## End GetEventErrors
## Begin DownTheRabbitHole
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

        $Values = $children | Where-Object { -not [string]::IsNullOrEmpty($_.PolicyState) }
        If ($Values) {
            ForEach ($Value in $Values) {
                If ($Value.PolicyState -eq "Delete") {
                    Write-Verbose "SETTING DELETE [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)]"
                    If ([string]::IsNullOrEmpty($_.Valuename)) {
                        Write-Warning "EMPTY VALUENAME, POTENTIAL SETTING FAILURE, CHECK MANUALLY [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)]"
                        Set-GPRegistryValue -Disable -Name $DestinationGPO -Key $Value.FullKeyPath -Verbose | Out-Null
                    }
                    Else {

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
                }
                Else {
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
                
        $subKeys = $children | Where-Object { [string]::IsNullOrEmpty($_.PolicyState) } | Select-Object -ExpandProperty FullKeyPath
        if ($subKeys) {
            DownTheRabbitHole -rootPaths $subKeys -SourceGPO $SourceGPOSingle -DestinationGPO $DestinationGPO -Verbose
        }
    }
}
## End DownTheRabbitHole
## Begin Launch-AzurePortal
Function Launch-AzurePortal { Invoke-Item "https://portal.azure.com/" -Credential (Get-Credential) }
## Begin Launch-ExchangeOnline
Function Launch-ExchangeOnline { Invoke-Item "https://outlook.office365.com/ecp/" }
## Begin Launch-InternetExplorer
Function Launch-InternetExplorer { & 'C:\Program Files (x86)\Internet Explorer\iexplore.exe' "about:blank" }
## Begin Launch-Office365Admin
Function Launch-Office365Admin { Invoke-Item "https://portal.office.com" -Credential (Get-Credential) }
## Begin Lock-Computer
Function Lock-Computer {
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
## End Lock-Computer
## Begin Reload-Profile
Function Reload-Profile {
    @(
        $Profile.AllUsersAllHosts,
        $Profile.AllUsersCurrentHost,
        $Profile.CurrentUserAllHosts,
        $Profile.CurrentUserCurrentHost
    ) | % {
        if (Test-Path $_) {
            Write-Verbose "Running $_"
            . $_
        }
    }    
}
## End Reload-Profile
## Begin Get-LastBoot
Function Get-LastBoot {
    <# 
  .SYNOPSIS 
  Retrieve last restart time for specified workstation(s) 

  .EXAMPLE 
  Get-LastBoot Computer123456 

  .EXAMPLE 
  Get-LastBoot 123456 
  #> 
    param([Parameter(Mandatory = $true)]
        [string[]] $ComputerName)
    if (($computername.length -eq 6)) {
        [int32] $dummy_output = $null;

        if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
            $computername = "Computer" + $computername.Replace("Computer", "")
        }	
    }

    $i = 0
    $j = 0

    foreach ($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

 
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{
            "Computer Name" = $Computer
            "Last Reboot"   = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}
## End Get-LastBoot
## Begin global:Get-LoggedOnUser
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
        $Choice = $host.ui.PromptForChoice($PromptTitle, $PromptText, $Options, 0)
        If ($Choice -eq 1) { break }
    }

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
    filter Invoke-Ping { (New-Object System.Net.NetworkInformation.Ping).Send($_, 10) }

    foreach ($comp in $Computers) { 
        # Check if computer needs to be Pinged first or not, and if Yes - see if responds to ping    
        switch ($DoNotPingComputer) {
            $false { $ProceedToCheck = ($Comp.Properties.dnshostname | Invoke-Ping).status -eq "Success" }
            $true { $ProceedToCheck = $true }
        }
     
        if ($ProceedToCheck) {   
            $user = gwmi win32_computersystem -ComputerName $Comp.Properties.dnshostname | select -ExpandProperty username1
            # If wmi query returned empty results - try querying with QUSER for active console session 
            if ($user -eq $null) {
                $user = quser /SERVER:$($Comp.Properties.dnshostname) | select-string active | % { $_.toString().split(" ")[1].Trim() }
            } 
            ## End global:Get-LoggedOnUser

            # Check if logged on user is a Direct member of Local Administrators group
            if ($user -eq $null) { $user = "No User logged On interactively" } 
            else { # Check if local admin
                # Note: locally can be checkd as- [Security.Principal.WindowsIdentity]::GetCurrent().IsInRole([Security.PrincipaltInRole] "Administrator")        
                $group = [ADSI]"WinNT://$($Comp.Properties.dnshostname)/administrators,group"
                $member = @($group.psbase.invoke("Members"))      
                $usersInGroup = $member | ForEach-Object { ([ADSI]$_).InvokeGet("Name") } 
                foreach ($GroupEntry in $usersInGroup) 
                { if ($GroupEntry -eq $user) { $AdminRole = $true } }
            }
            if ($AdminRole -ne $true -and $user -ne $null) { $AdminRole = $false } # if not admin, set to false     
            if ($ShowResultsToScreen) { write-host "$($Comp.Properties.dnshostname)`t$user`t$AdminRole" }
            $report += "$($Comp.Properties.dnshostname)`t$user`t$AdminRole"
            $user = $null
            $adminRole = $null
            $group = $null
            $member = $null
            $usersInGroup = $null
        } 
        else {
            # computer didn't respond to ping     
            $report += $($Comp.Properties.dnshostname) + "`tdidn't respond to ping - possibly Offile or Firewall issue"; $OfflineComputers += $($comp.properties.name)
            if ($ShowResultsToScreen) { Write-Warning "$($Comp.Properties.dnshostname)`tdidn't respond to ping - possibly  Offile or Port issue" }
        }
    }
    $report | Out-File $File 

    # Wrap up
    Write-Host "`nCompleted checking $($Computers.Count) hosts.`n" -ForegroundColor Green

    # check for offline computers, if encountered
    If ($OfflineComputers -ne $null) {
        # If there were offline / Non-responsive computers
        $OfflineComputers | Out-File "$ENV:Temp\NonRespondingComputers.txt"
        Write-Warning "Total of $($OfflineComputers.count) computers didn't respond to Ping.`nNon-Responding computers where saved into $($ENV:Temp)\NonRespondingComputers.txt." 
    }

    Write-Host "The full report was saved to $File" -ForegroundColor Cyan
    # Set back the system's current Error Action Preference
    $ErrorActionPreference = $CurrentEAP
}
## End global:Get-LoggedOnUser
#End Get-LoggedOnUser
## Begin Get-HotFixes
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
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = '')

    if (($computername.length -eq 6)) {
        [int32] $dummy_output = $null;

        if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
            $computername = "Computer" + $computername.Replace("Computer", "")
        }	
    }
    ## End Get-HotFixes

    $Stamp = (Get-Date -Format G) + ":"
    $ComputerArray = @()

    ## Begin HotFix
    Function HotFix {

        $i = 0
        $j = 0

        foreach ($computer in $ComputerArray) {

            Write-Progress -Activity "Retrieving HotFix Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

            Get-HotFix -Computername $computer 
        }    
    }
    ## End HotFix

    foreach ($computer in $ComputerName) {	     
        If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
            $ComputerArray += $Computer
        }	
    }
    ## End HotFix

    $HotFix = Get-HotFix
    $DocPath = [environment]::getfolderpath("mydocuments") + "\HotFix-Report.csv"

    Switch ($CheckBox.IsChecked) {
        $true { $HotFix | Export-Csv $DocPath -NoTypeInformation -Force; }
        default { $HotFix | Out-GridView -Title "HotFix Report"; }
    }

    if ($CheckBox.IsChecked -eq $true) {
        Try { 
            $listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
        } 

        Catch {
            #Do Nothing 
        }
    }
	
    else {
        Try {
            $listBox.Items.Add("$stamp HotFixes output processed!`n")
        } 
        Catch {
            #Do Nothing 
        }
    }
}
## End Get-HotFixes
## Begin CheckProcess
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
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = '')
    if (($computername.length -eq 6)) {
        [int32] $dummy_output = $null;

        if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
            $computername = "Computer" + $computername.Replace("Computer", "")
        }	
    }

    $Stamp = (Get-Date -Format G) + ":"
    $ComputerArray = @()

    ## Begin ChkProcess
    Function ChkProcess {

        $i = 0
        $j = 0

        foreach ($computer in $ComputerArray) {

            Write-Progress -Activity "Retrieving System Processes..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

            $getProcess = Get-Process -ComputerName $computer

            foreach ($Process in $getProcess) {
                
                [pscustomobject]@{
                    "Computer Name" = $computer
                    "Process Name"  = $Process.ProcessName
                    PID             = '{0:f0}' -f $Process.ID
                    Company         = $Process.Company
                    "CPU(s)"        = $Process.CPU
                    Description     = $Process.Description
                }           
            }
        } 
    }
    ## End ChkProcess
	
    foreach ($computer in $ComputerName) {	     
        If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
            $ComputerArray += $Computer
        }	
    }
    ## End ChkProcess
    $chkProcess = ChkProcess | Sort "Computer Name" | Select "Computer Name", "Process Name", PID, Company, "CPU(s)", Description
    $DocPath = [environment]::getfolderpath("mydocuments") + "\Process-Report.csv"

    Switch ($CheckBox.IsChecked) {
        $true { $chkProcess | Export-Csv $DocPath -NoTypeInformation -Force; }
        default { $chkProcess | Out-GridView -Title "Processes"; }
    }

    if ($CheckBox.IsChecked -eq $true) {
        Try { 
            $listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
        } 

        Catch {
            #Do Nothing 
        }
    }
	
    else {
        Try {
            $listBox.Items.Add("$stamp Check Process output processed!`n")
        } 
        Catch {
            #Do Nothing 
        }
    }
    
}
## End CheckProcess
## Begin Update-Sysinternals
Function Update-Sysinternals {
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
    [CmdletBinding()]
    param (
        # Path to the directory were sysinternals tools will be downloaded to 
        [Parameter(Mandatory = $true)]      
        [string]
        $Path 
    )
    
    begin {
        if (-not (Test-Path -Path $Path)) {
            Throw "The Path $_ does not exist"
        }
        else {
            $true
        }
        
        $uri = 'https://live.sysinternals.com/'
        $sysToolsPage = Invoke-WebRequest -Uri $uri
            
    }
    
    process {
        # create dir if it doesn't exist    
       
        Set-Location -Path $Path

        $sysTools = $sysToolsPage.Links.innerHTML | Where-Object -FilterScript { $_ -like "*.exe" -or $_ -like "*.chm" } 

        foreach ($sysTool in $sysTools) {
            Invoke-WebRequest -Uri "$uri/$sysTool" -OutFile $sysTool
        }
    } #process
}
## End Update-Sysinternals
## Begin Get-LocalAdmin
Function Get-LocalAdmin {
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    # Get local administrators
    $admins = Get-WmiObject -Class Win32_GroupUser -ComputerName $ComputerName |
              Where-Object { $_.GroupComponent -like '*"Administrators"' }

    # Extract and format user names
    $admins | ForEach-Object {
        if ($_.PartComponent -match 'Domain\=(.+?), Name\=(.+)$') {
            "$($matches[1].Trim('"'))\$($matches[2].Trim('"'))"
        }
    }
}
## End Get-LocalAdmin

## Begin findfile
Function findfile($name) {
    ls -recurse -filter "*${name}*" -ErrorAction SilentlyContinue | foreach {
        $place_path = $_.directory
        echo "${place_path}\${_}"
    }
}
## End findfile
## Begin rm-rf
Function rm-rf($item) { Remove-Item $item -Recurse -Force }
## Begin sudo
Function sudo() {
    Invoke-Elevated @args
}
## End sudo
## Begin PSsed
Function PSsed($file, $find, $replace) {
	(Get-Content $file).replace("$find", $replace) | Set-Content $file
}
## End PSsed
## Begin PSsed-recursive
Function Set-ContentRecursive($filePattern, $find, $replace) {
    $files = Get-ChildItem . "$filePattern" -Recurse # -Exclude
    foreach ($file in $files) {
		(Get-Content $file.PSPath) |
        Foreach-Object { $_ -replace "$find", "$replace" } |
        Set-Content $file.PSPath
    }
}
## End PSsed-recursive
## Begin PSgrep
Function PSgrep {

    [CmdletBinding()]
    Param(
    
        # source file to grep
        [Parameter(Mandatory = $true)]
        [string]$SourceFileName, 

        # string to search for
        [Parameter(Mandatory = $true)]
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
            if ($_ -match $String) {
                $count ++
                if (!($OutputFile)) {
                    write-host $_
                }
                else {
                    $_ | Out-File -FilePath ".\$($OutputFile)" -Append -Force
                }

            }

        }

    }

    Write-Host "$($Count) matches found"
}
## End PSgrep
## Begin which
Function which($name) {
    Get-Command $name | Select-Object -ExpandProperty Definition
}
## End which
## Begin cut
Function cut() {
    foreach ($part in $input) {
        $line = $part.ToString();
        $MaxLength = [System.Math]::Min(200, $line.Length)
        $line.subString(0, $MaxLength)
    }
}
## End cut
## Begin Search-AllTextFiles
Function Search-AllTextFiles {
    param(
        [parameter(Mandatory = $true, position = 0)]$Pattern, 
        [switch]$CaseSensitive,
        [switch]$SimpleMatch
    );

    Get-ChildItem . * -Recurse -Exclude ('*.dll', '*.pdf', '*.pdb', '*.zip', '*.exe', '*.jpg', '*.gif', '*.png', '*.ico', '*.svg', '*.bmp', '*.psd', '*.cache', '*.doc', '*.docx', '*.xls', '*.xlsx', '*.dat', '*.mdf', '*.nupkg', '*.snk', '*.ttf', '*.eot', '*.woff', '*.tdf', '*.gen', '*.cfs', '*.map', '*.min.js', '*.data') | Select-String -Pattern:$pattern -SimpleMatch:$SimpleMatch -CaseSensitive:$CaseSensitive
}
## End Search-AllTextFiles
## Begin AddTo-7zip
Function AddTo-7zip($zipFileName) {
    BEGIN {
        #$7zip = "$($env:ProgramFiles)\7-zip\7z.exe"
        $7zip = Find-Program "\7-zip\7z.exe"
        if (!([System.IO.File]::Exists($7zip))) {
            throw "7zip not found";
        }
    }
    PROCESS {
        & $7zip a -tzip $zipFileName $_
    }
    END {
    }
}
## End AddTo-7zip
## Begin GoGo-PSExch
Function GoGo-PSExch {

    param(
        [Parameter( Mandatory = $false)]
        [string]$URL = "MWTEXCH01"
    )
    
    #$Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

    $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos #-Credential #$Credentials

    Import-PSSession $ExOPSession -AllowClobber

}
## End GoGo-PSExch

## Begin Get-FileAttribute
Function Get-FileAttribute {
    param($file, $attribute)
    $val = [System.IO.FileAttributes]$attribute;
    if ((Get-ChildItem $file -force).Attributes -band $val -eq $val) { $true; } else { $false; }
} 
## End Get-FileAttribute
## Begin Set-FileAttribute
Function Set-FileAttribute {
    param($file, $attribute)
    $file = (Get-ChildItem $file -force);
    $file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]$attribute).value__;
    if ($?) { $true; } else { $false; }
} 
## End Set-FileAttribute
## Begin LastBoot
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

        [Parameter(Mandatory = $true)]
        [String[]]$ComputerName,

        $i = 0,
        $j = 0
    )

    foreach ($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)
 
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{

            ComputerName = $Computer
            LastReboot   = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}#End LastBoot
## End LastBoot
## Begin SYSinfo
Function SYSinfo {
    <# 
.SYNOPSIS 
  Retrieve basic system information for specified workstation(s) 

.EXAMPLE 
  SYS Computer123456 
#> 

    param(

        [Parameter(Mandatory = $true)]
        [string[]] $ComputerName,
    
        $i = 0,
        $j = 0
    )

    $Stamp = (Get-Date -Format G) + ":"

    Function Systeminformation {
	
        foreach ($Computer in $ComputerName) {

            if (!([String]::IsNullOrWhiteSpace($Computer))) {

                if (Test-Connection -Quiet -Count 1 -Computer $Computer) {

                    Write-Progress -Activity "Getting Sytem Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

                    Start-Job -ScriptBlock { param($Computer) 

                        #Gather specified workstation information; CimInstance only works on 64-bit
                        $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer
                        $computerBIOS = Get-CimInstance CIM_BIOSElement -Computer $Computer
                        $computerOS = Get-CimInstance CIM_OperatingSystem -Computer $Computer
                        $computerCPU = Get-CimInstance CIM_Processor -Computer $Computer
                        $computerHDD = Get-CimInstance Win32_LogicalDisk -Computer $Computer -Filter "DeviceID = 'C:'"
    
                        [PSCustomObject] @{

                            ComputerName    = $computerSystem.Name
                            LastReboot      = $computerOS.LastBootUpTime
                            OperatingSystem = $computerOS.OSArchitecture + " " + $computerOS.caption
                            Model           = $computerSystem.Model
                            RAM             = "{0:N2}" -f [int]($computerSystem.TotalPhysicalMemory / 1GB) + "GB"
                            DiskCapacity    = "{0:N2}" -f ($computerHDD.Size / 1GB) + "GB"
                            TotalDiskSpace  = "{0:P2}" -f ($computerHDD.FreeSpace / $computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace / 1GB) + "GB)"
                            CurrentUser     = $computerSystem.UserName
                        }
                    } -ArgumentList $Computer
                }

                else {

                    Start-Job -ScriptBlock { param($Computer)  
                     
                        [PSCustomObject] @{

                            ComputerName    = $Computer
                            LastReboot      = "Unable to PING."
                            OperatingSystem = "$Null"
                            Model           = "$Null"
                            RAM             = "$Null"
                            DiskCapacity    = "$Null"
                            TotalDiskSpace  = "$Null"
                            CurrentUser     = "$Null"
                        }
                    } -ArgumentList $Computer                       
                }
            }

            else {
                 
                Start-Job -ScriptBlock { param($Computer)  
                     
                    [PSCustomObject] @{

                        ComputerName    = "Value is null."
                        LastReboot      = "$Null"
                        OperatingSystem = "$Null"
                        Model           = "$Null"
                        RAM             = "$Null"
                        DiskCapacity    = "$Null"
                        TotalDiskSpace  = "$Null"
                        CurrentUser     = "$Null"
                    }
                } -ArgumentList $Computer
            }
        } 
    }

    $SystemInformation = SystemInformation | Receive-Job -Wait | Select-Object ComputerName, CurrentUser, OperatingSystem, Model, RAM, DiskCapacity, TotalDiskSpace, LastReboot
    $DocPath = [environment]::getfolderpath("mydocuments") + "\SystemInformation-Report.csv"

    Switch ($CheckBox.IsChecked) {

        $true { 
            
            $SystemInformation | Export-Csv $DocPath -NoTypeInformation -Force 
        }

        default { 
            
            $SystemInformation | Out-GridView -Title "System Information"
        }
    }

    if ($CheckBox.IsChecked -eq $true) {

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
## End SYSinfo
## Begin NetMSG
Function NetMSG {
    <# 
.SYNOPSIS 
    Generate a pop-up window on specified workstation(s) with desired message 

.EXAMPLE 
    NetMSG Computer123456 
#> 
	
    param(

        [Parameter(Mandatory = $true)]
        [String[]] $ComputerName,

        [Parameter(Mandatory = $true, HelpMessage = 'Enter desired message')]
        [String]$MyMessage,

        [String]$User = [Environment]::UserName,

        [String]$UserJob = (Get-ADUser $User -Property Title).Title,
    
        [String]$CallBack = "$User | 5-2444 | $UserJob",

        $i = 0,
        $j = 0
    )

    Function SendMessage {

        foreach ($Computer in $ComputerName) {

            Write-Progress -Activity "Sending messages..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)         

            #Invoke local MSG command on specified workstation - will generate pop-up message for any user logged onto that workstation - *Also shows on Login screen, stays there for 100,000 seconds or until interacted with
            Invoke-Command -ComputerName $Computer { param($MyMessage, $CallBack, $User, $UserJob)
 
                MSG /time:100000 * /v "$MyMessage {$CallBack}"
            } -ArgumentList $MyMessage, $CallBack, $User, $UserJob -AsJob
        }
    }

    SendMessage | Wait-Job | Remove-Job

}#End NetMSG
## End NetMSG
## Begin InstallApplication
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

        [Parameter(Mandatory = $true, HelpMessage = "Enter Computername(s)")]
        [String[]]$Computername,

        [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter installer path(s)")]
        [String[]]$Path = $null,

        [Parameter(ValueFromPipeline = $true, HelpMessage = 'Enter remote destination: C$\Directory')]
        $Destination = "C$\TempApplications"
    )

    if ($null -eq $Path) {

        Add-Type -AssemblyName System.Windows.Forms

        $Dialog = New-Object System.Windows.Forms.OpenFileDialog
        $Dialog.InitialDirectory = "\\lasfs03\Software\Current Version\Deploy"
        $Dialog.Title = "Select Installation File(s)"
        $Dialog.Filter = "Installation Files (*.exe,*.msi,*.msp)|*.exe; *.msi; *.msp"        
        $Dialog.Multiselect = $true
        $Result = $Dialog.ShowDialog()

        if ($Result -eq 'OK') {

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
        foreach ($Computer in $Computername) {

            #If $Computer IS NOT null or only whitespace
            if (!([string]::IsNullOrWhiteSpace($Computer))) {

                #Test-Connection to $Computer
                if (Test-Connection -Quiet -Count 1 $Computer) {                                               
                     
                    #Create job on localhost
                    Start-Job { param($Computer, $Path, $Destination)

                        foreach ($P in $Path) {
                            
                            #Static Temp location
                            $TempDir = "\\$Computer\$Destination"

                            #Create $TempDir directory
                            if (!(Test-Path $TempDir)) {

                                New-Item -Type Directory $TempDir | Out-Null
                            }
                     
                            #Retrieve Leaf object from $Path
                            $FileName = (Split-Path -Path $P -Leaf)

                            #New Executable Path
                            $Executable = "C:\$(Split-Path -Path $Destination -Leaf)\$FileName"

                            #Copy needed installer files to remote machine
                            Copy-Item -Path $P -Destination $TempDir

                            #Install .EXE
                            if ($FileName -like "*.exe") {

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
                            elseif ($FileName -like "*.msi") {

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
                            elseif ($FileName -like "*.msp") { 
                                                                       
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
## End InstallApplication
## Begin Get-Icon
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
        [parameter(ValueFromPipelineByPropertyName = $True)]
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
            $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path) | 
            Add-Member -MemberType NoteProperty -Name FullName -Value $Path -PassThru
            If ($PSBoundParameters.ContainsKey('ToBytes')) {
                Write-Verbose "Retrieving bytes"
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $MemoryStream.ToArray()   
                $MemoryStream.Flush()  
                $MemoryStream.Dispose()           
            }
            ElseIf ($PSBoundParameters.ContainsKey('ToBitmap')) {
                $Icon.ToBitMap()
            }
            ElseIf ($PSBoundParameters.ContainsKey('ToBase64')) {
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $Bytes = $MemoryStream.ToArray()   
                $MemoryStream.Flush() 
                $MemoryStream.Dispose()
                [convert]::ToBase64String($Bytes)
            }
            Else {
                $Icon
            }
        }
        Else {
            Write-Warning "$Path does not exist!"
            Continue
        }
    }
}
## End Get-Icon
## Begin Get-MappedDrive
Function Get-MappedDrive {
    param (
        [string]$computername = "localhost"
    )
    Get-WmiObject -Class Win32_MappedLogicalDisk -ComputerName $computername | 
    Format-List DeviceId, VolumeName, SessionID, Size, FreeSpace, ProviderName
}
# End Get Mapped Drive
## Begin LayZ
Function LayZ {
    C:\LazyWinAdmin\LazyWinAdmin\LazyWinAdmin.ps1
}
# End LayZ
## Begin Get-UserLastLogonTime
Function Get-UserLastLogonTime {

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
    Begin {
        #List of users that are present on all PCs, these won't output unless specified to do so.
        $CommonUsers = ("NetworkService", "LocalService", "systemprofile")
    }
    
    #Process Pipeline
    Process {
        #ping the machine before trying to do anything
        if (Test-Connection $ComputerName -Count 2 -Quiet) {
            #try to get the OS version of the computer
            try { $OS = (gwmi  -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction stop).caption }
            catch {
                #had an error getting the WMI-object
                return New-Object psobject -Property @{
                    User        = "Error getting WMIObject Win32_OperatingSystem"
                    LastUseTime = get-date 0
                }
            }
            #make sure the OS retrieved is either Windows 7 or Windows 10 as this Function has not been set to work on other operating systems
            if ($OS.contains("Windows 10") -or $OS.Contains("Windows 7")) {
                try {
                    #try to get the WMiObject win32_UserProfile
                    $UserObjects = Get-WmiObject -ComputerName $ComputerName -Class Win32_userProfile -Property LocalPath, LastUseTime -ErrorAction Stop
                    $users = @()
                    #loop that handles all the data that came back from the WMI objects
                    forEach ($UserObject in $UserObjects) {
                        #extract the username from the local path, first find where the last slash is after, get everything after that into its own string
                        $i = $UserObject.localPath.LastIndexOf("\") + 1
                        $tempUserString = ""
                        while ($null -ne $UserObject.localPath.toCharArray()[$i]) {
                            $tempUserString += $UserObject.localPath.toCharArray()[$i]
                            $i++
                        }
                        #if list common users is turned on, skip this block
                        if (!$listCommonUsers) {
                            #check if the user extracted from the local path is a common user
                            $isCommonUser = $false
                            forEach ($userName in $CommonUsers) { 
                                if ($userName -eq $tempUserString) {
                                    $isCommonUser = $true
                                    break
                                }
                            }
                        }
                        #if the user is one of the users specified in the common users, skip it unless otherwise specified
                        if ($isCommonUser) { continue }
                        #check to see if the user has a timestamp for there last logon 
                        if ($null -ne $UserObject.LastUseTime) {
                            #This converts the string timestamp to a DateTime object
                            $TempUserLastUseTime = ([WMI] '').ConvertToDateTime($userObject.LastUseTime)
                        }
                        #the user had a local path but no timestamp was found for last logon
                        else { $TempUserLastUseTime = Get-Date 0 }
                        #add user to array
                        $users += New-Object -TypeName psobject -Property @{
                            User        = $tempUserString
                            LastUseTime = $TempUserLastUseTime
                        }
                    }
                }
                catch {
                    #error trying to retrieve WMI obeject win32_userProfile
                    return New-Object psobject -Property @{
                        User        = "Error getting WMIObject Win32_userProfile"
                        LastUseTime = get-date 0
                    }
                }
            }
            else {
                #OS version was not compatible
                return New-Object psobject -Property @{
                    User        = "Operating system $OS is not compatible with this Function."
                    LastUseTime = get-date 0
                }
            }
        }
        else {
            #Computer was not pingable
            return New-Object psobject -Property @{
                User        = "Can't Ping"
                LastUseTime = get-date 0
            }
        }

        #check to see if any users came out of the main Function
        if ($users.count -eq 0) {
            $users += New-Object -TypeName psobject -Property @{
                User        = "NoLoggedOnUsers"
                LastUseTime = get-date 0
            }
        }
        #sort the user array by the last time they logged in
        else { $users = $users | Sort-Object -Property LastUseTime -Descending }
        #main output block
        #if List all users was chosen, output the full list of users found
        if ($ListAllUsers) { return $users }
        #if get last user was chosen, output the last user to log on the computer
        elseif ($GetLastUser) { return ($users[0]) }
        else {
            #see if the user specified ever logged on
            ForEach ($Username in $users) {
                if ($Username.User -eq $user) { return ($Username) }            
            }

            #user did not log on
            return New-Object psobject -Property @{
                User        = "$user"
                LastUseTime = get-date 0
            }
        }
    }
    #End Pipeline
    End { Write-Verbose "Function get-UserLastLogonTime is complete" }
}
## End Get-UserLastLogonTime
## Begin Unblock
Function Unblock ($path) { 

    Get-ChildItem "$path" -Recurse | Unblock-File

}
## End Unblock
## Begin Get-RemoteSysInfo
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

        [Parameter(Mandatory = $true)]
        [string[]] $ComputerName
    )

    $Stamp = (Get-Date -Format G) + ":"
    $ComputerArray = @()

    $i = 0
    $j = 0

    ## Begin Systeminformation
    Function Systeminformation {
	
        foreach ($Computer in $ComputerName) {

            if (!([String]::IsNullOrWhiteSpace($Computer))) {

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

                            "Computer Name"    = $computerSystem.Name
                            "Last Reboot"      = $computerOS.LastBootUpTime
                            "Operating System" = $computerOS.OSArchitecture + " " + $computerOS.caption
                            Model              = $computerSystem.Model
                            RAM                = "{0:N2}" -f [int]($computerSystem.TotalPhysicalMemory / 1GB) + "GB"
                            "Disk Capacity"    = "{0:N2}" -f ($computerHDD.Size / 1GB) + "GB"
                            "Total Disk Space" = "{0:P2}" -f ($computerHDD.FreeSpace / $computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace / 1GB) + "GB)"
                            "Current User"     = $computerSystem.UserName
                        }
                    } -ArgumentList $Computer
                }

                else {

                    Start-Job -ScriptBlock { param($Computer)  
                     
                        [pscustomobject]@{

                            "Computer Name"    = $Computer
                            "Last Reboot"      = "Unable to PING."
                            "Operating System" = "$Null"
                            Model              = "$Null"
                            RAM                = "$Null"
                            "Disk Capacity"    = "$Null"
                            "Total Disk Space" = "$Null"
                            "Current User"     = "$Null"
                        }
                    } -ArgumentList $Computer                       
                }
            }

            else {
                 
                Start-Job -ScriptBlock { param($Computer)  
                     
                    [pscustomobject]@{

                        "Computer Name"    = "Value is null."
                        "Last Reboot"      = "$Null"
                        "Operating System" = "$Null"
                        Model              = "$Null"
                        RAM                = "$Null"
                        "Disk Capacity"    = "$Null"
                        "Total Disk Space" = "$Null"
                        "Current User"     = "$Null"
                    }
                } -ArgumentList $Computer
            }
        } 
    }
    ## End Systeminformation

    $SystemInformation = SystemInformation | Wait-Job | Receive-Job | Select-Object "Computer Name", "Current User", "Operating System", Model, RAM, "Disk Capacity", "Total Disk Space", "Last Reboot"
    $DocPath = [environment]::getfolderpath("mydocuments") + "\SystemInformation-Report.csv"

    Switch ($CheckBox.IsChecked) {
        $true { $SystemInformation | Export-Csv $DocPath -NoTypeInformation -Force; }
        default { $SystemInformation | Out-GridView -Title "System Information"; }
		
    }

    if ($CheckBox.IsChecked -eq $true) {

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
}
## End Get-RemoteSysInfo
## Begin Get-RemoteSoftWare
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
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = '')
    if (($computername.length -eq 6)) {
        [int32] $dummy_output = $null;

        if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
            $computername = "Computer" + $computername.Replace("Computer", "")
        }	
    }

    $Stamp = (Get-Date -Format G) + ":"
    $ComputerArray = @()

    ## Begin SoftwareCheck
    Function SoftwareCheck {

        $i = 0
        $j = 0

        foreach ($computer in $ComputerArray) {

            Write-Progress -Activity "Retrieving Software Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

            $keys = '', '\Wow6432Node'
            foreach ($key in $keys) {
                try {
                    $apps = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer).OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall").GetSubKeyNames()
                }
                catch {
                    continue
                }

                foreach ($app in $apps) {
                    $program = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer).OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall\$app")
                    $name = $program.GetValue('DisplayName')
                    if ($name -and $name -match $NameRegex) {
                        [pscustomobject]@{
                            "Computer Name"    = $computer
                            Software           = $name
                            Version            = $program.GetValue('DisplayVersion')
                            Publisher          = $program.GetValue('Publisher')
                            "Install Date"     = $program.GetValue('InstallDate')
                            "Uninstall String" = $program.GetValue('UninstallString')
                            Bits               = $(if ($key -eq '\Wow6432Node') { '64' } else { '32' })
                            Path               = $program.name
                        }
                    }
                }
            } 
        }
    }	
    ## End SoftwareCheck

    foreach ($computer in $ComputerName) {	     
        If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
            $ComputerArray += $Computer
        }	
    }
    ## End SoftwareCheck
    $SoftwareCheck = SoftwareCheck | Sort-Object "Computer Name" | Select-Object "Computer Name", Software, Version, Publisher, "Install Date", "Uninstall String", Bits, Path
    $DocPath = [environment]::getfolderpath("mydocuments") + "\Software-Report.csv"

    Switch ($CheckBox.IsChecked) {
        $true { $SoftwareCheck | Export-Csv $DocPath -NoTypeInformation -Force; }
        default { $SoftwareCheck | Out-GridView -Title "Software"; }
    }
		
    if ($CheckBox.IsChecked -eq $true) {
        Try { 
            $listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
        } 

        Catch {
            #Do Nothing 
        }
    }
	
    else {
        Try {
            $listBox.Items.Add("$stamp Software output processed!`n")
        } 
        Catch {
            #Do Nothing 
        }
    }
}
## End Get-RemoteSoftWare
## Begin Get-OfficeVersion
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
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
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

        $defaultDisplaySet = 'DisplayName', 'Version', 'ComputerName'

        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(�DefaultDisplayPropertySet�, [string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    }
    ## End Get-OfficeVersion

    process {

        $results = new-object PSObject[] 0;

        foreach ($computer in $ComputerName) {
            if ($Credentials) {
                $os = Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
            }
            else {
                $os = Get-WMIObject win32_operatingsystem -computername $computer
            }

            $osArchitecture = $os.OSArchitecture

            if ($Credentials) {
                $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
            }
            else {
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
                            $mainlangCulture = [globalization.cultureinfo]::GetCultures("allCultures") | Where-Object { $_.LCID -eq $mainLangId }
                            if ($mainlangCulture) {
                                $cltr.ClientCulture = $mainlangCulture.Name
                            }
                        }

                        [string]$officeLangPath = join-path  $path "Common\LanguageResources\InstalledUIs"
                        $langValues = $regProv.EnumValues($HKLM, $officeLangPath);
                        if ($langValues) {
                            foreach ($langValue in $langValues) {
                                [globalization.cultureinfo]::GetCultures("allCultures") | Where-Object { $_.LCID -eq $langValue } | Out-Null
                            } 
                        }

                        if ($virtualInstallPath) {

                        }
                        else {
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
                # Removed unused variable $keyList
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
                        Bitness = $buildType; ComputerName = $computer; ClickToRunUpdatesEnabled = $cltrUpdatedEnabled; ClickToRunUpdateUrl = $cltrUpdateUrl;
                        ClientCulture = $clientCulture 
                    }
                    $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                    $results += $object

                }
            }

        }

        $results = Get-Unique -InputObject $results 

        return $results;
    }
    ## End Get-OfficeVersion

}
## End Get-OfficeVersion
## Begin Get-OfficeVersion2
Function Get-OfficeVersion2 {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $Infile,
    
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string] $outfile
    )
    #$outfile = 'C:\temp\office.csv'
    #$infile = 'c:\temp\servers.txt'
    Begin {
    }
    Process {
        $office = @()
        $computers = Get-Content $infile
        $i = 0
        $count = $computers.count
        foreach ($computer in $computers) {
            $i++
            Write-Progress -Activity "Querying Computers" -Status "Computer: $i of $count " `
                -PercentComplete ($i / $count * 100)
            $info = @{}
            $version = 0
            try {
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer) 
                $reg.OpenSubKey('software\Microsoft\Office').GetSubKeyNames() | ForEach-Object {
                    if ($_ -match '(\d+)\.') {
                        if ([int]$matches[1] -gt $version) {
                            $version = $matches[1]
                        }
                    }    
                }
                if ($version) {
                    Write-Debug("$computer : found $version")
                    switch ($version) {
                        "7" { $officename = 'Office 97' }
                        "8" { $officename = 'Office 98' }
                        "9" { $officename = 'Office 2000' }
                        "10" { $officename = 'Office XP' }
                        "11" { $officename = 'Office 97' }
                        "12" { $officename = 'Office 2003' }
                        "13" { $officename = 'Office 2007' }
                        "14" { $officename = 'Office 2010' }
                        "15" { $officename = 'Office 2013' }
                        "16" { $officename = 'Office 2016' }
                        default { $officename = 'Unknown Version' }
                    }
    
                }
            }
            catch {
                $officename = 'Not Installed/Not Available'
            }
            $info.Computer = $computer
            $info.Name = $officename
            $info.version = $version

            $object = new-object -TypeName PSObject -Property $info
            $office += $object
        }
        $office | Select-Object computer, version, name | Export-Csv -NoTypeInformation -Path $outfile
    }
}
write-output ("Done")
## End Get-OfficeVersion2
## Begin Get-OutlookClientVersion
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
            $Headers = (Get-Content -Path $file -TotalCount 5 | Where-Object { $_ -like '#Fields*' }) -replace '#Fields: ' -split ','
                    
            Import-Csv -Header $Headers -Path $file |
            Where-Object { $_.operation -eq 'Connect' -and $_.'client-software' -eq 'outlook.exe' } |
            Select-Object -Unique -Property @{label = 'User'; expression = { $_.'client-name' -replace '^.*cn=' } },
            @{label = 'DN'; expression = { $_.'client-name' } },
            client-software,
            @{label = 'Version'; expression = { Get-MrOutlookVersion -OutlookBuild $_.'client-software-version' } },
            client-mode,
            client-ip,
            protocol
        }
    }
}
## End Get-OutlookClientVersion
## Begin Get-MSOutlookVersion
Function Get-MSOutlookVersion {
    param (
        [string]$OutlookBuild
    )
    switch ($OutlookBuild) {  
        { $_ -ge '16.0.4266.1001' } { 'Outlook 2016 4266.1001'; break }
        { $_ -ge '16.0.4522.1000' } { 'Outlook 2016 4522.1000'; break }
        { $_ -ge '16.0.4498.1000' } { 'Outlook 2016 4498.1000'; break }
        { $_ -ge '16.0.4229.1024' } { 'Outlook 2016 4229.1024'; break }              
        { $_ -ge '15.0.4569.1506' } { 'Outlook 2013 SP1'; break }
        { $_ -ge '15.0.4420.1017' } { 'Outlook 2013 RTM'; break }
        { $_ -ge '14.0.7015.1000' } { 'Outlook 2010 SP2'; break }
        { $_ -ge '14.0.6029.1000' } { 'Outlook 2010 SP1'; break }
        { $_ -ge '14.0.4763.1000' } { 'Outlook 2010 RTM'; break }
        { $_ -ge '12.0.6672.5000' } { 'Outlook 2007 SP3 U2013'; break }
        { $_ -ge '12.0.6423.1000' } { 'Outlook 2007 SP2'; break }
        { $_ -ge '12.0.6212.1000' } { 'Outlook 2007 SP1'; break }
        { $_ -ge '12.0.4518.1014' } { 'Outlook 2007 RTM'; break }
        Default { '$OutlookBuild' }
    }
}
## End Get-MSOutlookVersion
## Begin Get-UserLogon
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
                ## End Get-UserLogon
 
                $result += New-Object -TypeName PSCustomObject -Property $array
 
            }
            ## End Get-UserLogon
 
            else {
 
                $array = ([ordered]@{
                        'User'     = $b[0]
                        'Computer' = $Computer
                        'Date'     = $b[5]
                        'Time'     = $b[6..7] -join ' '
                    })
                ## End Get-UserLogon
 
                $result += New-Object -TypeName PSCustomObject -Property $array
 
            }
            ## End Get-UserLogon
        }
        ## End Get-UserLogon
    }
    ## End Get-UserLogon
 
    If ($OU) {
 
        $comp = Get-ADComputer -Filter * -SearchBase "$OU" -Properties operatingsystem
 
        $count = $comp.count
 
        If ($count -gt 20) {
 
            Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer"
 
        }
        ## End Get-UserLogon
 
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
                    ## End Get-UserLogon
 
                    $result += New-Object -TypeName PSCustomObject -Property $array
                }
                ## End Get-UserLogon
 
                else {
 
                    $array = ([ordered]@{
                            'User'     = $a[0]
                            'Computer' = $u.Name
                            'Date'     = $a[5]
                            'Time'     = $a[6..7] -join ' '
                        })
                    ## End Get-UserLogon
 
                    $result += New-Object -TypeName PSCustomObject -Property $array
                }
                ## End Get-UserLogon
 
            }
            ## End Get-UserLogon
 
        }
        ## End Get-UserLogon
 
    }
    ## End Get-UserLogon
 
    If ($All) {
 
        $comp = Get-ADComputer -Filter * -Properties operatingsystem
 
        $count = $comp.count
 
        If ($count -gt 20) {
 
            Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer ..."
 
        }
        ## End Get-UserLogon
 
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
                    ## End Get-UserLogon
 
                    $result += New-Object -TypeName PSCustomObject -Property $array
 
                }
                ## End Get-UserLogon
 
                else {
 
                    $array = ([ordered]@{
                            'User'     = $a[0]
                            'Computer' = $u.Name
                            'Date'     = $a[5]
                            'Time'     = $a[6..7] -join ' '
                        })
                    ## End Get-UserLogon
 
                    $result += New-Object -TypeName PSCustomObject -Property $array
 
                }
                ## End Get-UserLogon
 
            }
            ## End Get-UserLogon
 
        }
        ## End Get-UserLogon
    }
    ## End Get-UserLogon
    Write-Output $result
}
## End Get-UserLogon
## Begin New-IsoFile
Function New-IsoFile {  
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
  
    [CmdletBinding(DefaultParameterSetName = 'Source')]Param( 
        [parameter(Position = 1, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Source')]$Source,  
        [parameter(Position = 2)][string]$Path = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",  
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })][string]$BootFile = $null, 
        [ValidateSet('CDR', 'CDRW', 'DVDRAM', 'DVDPLUSR', 'DVDPLUSRW', 'DVDPLUSR_DUALLAYER', 'DVDDASHR', 'DVDDASHRW', 'DVDDASHR_DUALLAYER', 'DISK', 'DVDPLUSRW_DUALLAYER', 'BDR', 'BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER', 
        [string]$Title = (Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
        [switch]$Force, 
        [parameter(ParameterSetName = 'Clipboard')][switch]$FromClipboard 
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
## End New-IsoFile
'@  
        } 
  
        if ($BootFile) { 
            if ('BDR', 'BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
      ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type = 1 }).Open()  # adFileTypeBinary 
            $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
      ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
        } 
 
        $MediaType = @('UNKNOWN', 'CDROM', 'CDR', 'CDRW', 'DVDROM', 'DVDRAM', 'DVDPLUSR', 'DVDPLUSRW', 'DVDPLUSR_DUALLAYER', 'DVDDASHR', 'DVDDASHRW', 'DVDDASHR_DUALLAYER', 'DISK', 'DVDPLUSRW_DUALLAYER', 'HDDVDROM', 'HDDVDR', 'HDDVDRAM', 'BDROM', 'BDR', 'BDRE') 
 
        Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))" 
    ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName = $Title }).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 
  
        if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break } 
    }  
 
    Process { 
        if ($FromClipboard) { 
            if ($PSVersionTable.PSVersion.Major -lt 5) { Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break } 
            $Source = Get-Clipboard -Format FileDropList 
        } 
 
        foreach ($item in $Source) { 
            if ($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) { 
                $item = Get-Item -LiteralPath $item 
            } 
 
            if ($item) { 
                Write-Verbose -Message "Adding item to the target image: $($item.FullName)" 
                try { $Image.Root.AddTree($item.FullName, $true) } catch { Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') } 
            } 
        } 
    } 
 
    End {  
        if ($Boot) { $Image.BootImageOptions = $Boot }  
        $Result = $Image.CreateResultImage()  
        [ISOFile]::Create($Target.FullName, $Result.ImageStream, $Result.BlockSize, $Result.TotalBlocks) 
        Write-Verbose -Message "Target image ($($Target.FullName)) has been created" 
        $Target 
    } 
} 
## End New-IsoFile
## Begin Set-FileTime
Function Set-FileTime {
    param(
        [string[]]$paths,
        [bool]$only_modification = $false,
        [bool]$only_access = $false
    );

    begin {
        Function updateFileSystemInfo([System.IO.FileSystemInfo]$fsInfo) {
            $datetime = get-date
            if ( $only_access ) {
                $fsInfo.LastAccessTime = $datetime
            }
            elseif ( $only_modification ) {
                $fsInfo.LastWriteTime = $datetime
            }
            else {
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
## End Set-FileTime
## Begin Get-PendingUpdate
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
                    $updatesession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session", $computer)) 
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
                    For ($i = 0; $i -lt $Count; $i++) { 
                        #Create object holding update 
                        $Update = $searchresult.Updates.Item($i)
                        [pscustomobject]@{
                            Computername     = $Computer
                            Title            = $Update.Title
                            KB               = $($Update.KBArticleIDs)
                            SecurityBulletin = $($Update.SecurityBulletinIDs)
                            MsrcSeverity     = $Update.MsrcSeverity
                            IsDownloaded     = $Update.IsDownloaded
                            Url              = $($Update.MoreInfoUrls)
                            Categories       = ($Update.Categories | Select-Object -ExpandProperty Name)
                            BundledUpdates   = @($Update.BundledUpdates) | ForEach-Object {
                                [pscustomobject]@{
                                    Title       = $_.Title
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
## End Get-PendingUpdate
## Begin Get-NetworkLevelAuthentication
Function Get-NetworkLevelAuthentication {
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
        [System.Management.Automation.PSCredential]$Credential
    )#Param
    BEGIN {
        TRY {
            IF (-not (Get-Module -Name CimCmdlets)) {
                Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
                Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
            }
        }
        CATCH {
            IF ($ErrorBeginCimCmdlets) {
                Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
            }
        }
    }#BEGIN
	
    PROCESS {
        FOREACH ($Computer in $ComputerName) {
            TRY {
                # Building Splatting for CIM Sessions
                $CIMSessionParams = @{
                    ComputerName  = $Computer
                    ErrorAction   = 'Stop'
                    ErrorVariable = 'ProcessError'
                }
				
                # Add Credential if specified when calling the Function
                IF ($PSBoundParameters['Credential']) {
                    $CIMSessionParams.credential = $Credential
                }
				
                # Connectivity Test
                Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
                Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
                # CIM/WMI Connection
                #  WsMAN
                IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0') {
                    Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
                    $CimSession = New-CimSession @CIMSessionParams
                    $CimProtocol = $CimSession.protocol
                    Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
                }
				
                # DCOM
                ELSE {
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
                    'ComputerName'     = $NLAinfo.PSComputerName
                    'NLAEnabled'       = $NLAinfo.UserAuthenticationRequired -as [bool]
                    'TerminalName'     = $NLAinfo.TerminalName
                    'TerminalProtocol' = $NLAinfo.TerminalProtocol
                    'Transport'        = $NLAinfo.transport
                }
            }
			
            CATCH {
                Write-Warning -Message "PROCESS - Error on $Computer"
                $_.Exception.Message
                if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
                if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
            }#CATCH
        } # FOREACH
    }#PROCESS
    END {
		
        if ($CimSession) {
            Write-Verbose -Message "END - Close CIM Session(s)"
            Remove-CimSession $CimSession
        }
        Write-Verbose -Message "END - Script is completed"
    }
}
## End Get-NetworkLevelAuthentication
## Begin Set-NetworkLevelAuthentication
Function Set-NetworkLevelAuthentication {
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
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [String[]]$ComputerName = $env:ComputerName,
		
        [Parameter(Mandatory)]
        [Bool]$EnableNLA,
		
        [Alias("RunAs")]
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential
    )#Param
    BEGIN {
        TRY {
            IF (-not (Get-Module -Name CimCmdlets)) {
                Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
                Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
				
            }
        }
        CATCH {
            IF ($ErrorBeginCimCmdlets) {
                Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
            }
        }
    }#BEGIN
	
    PROCESS {
        FOREACH ($Computer in $ComputerName) {
            TRY {
                # Building Splatting for CIM Sessions
                $CIMSessionParams = @{
                    ComputerName  = $Computer
                    ErrorAction   = 'Stop'
                    ErrorVariable = 'ProcessError'
                }
				
                # Add Credential if specified when calling the Function
                IF ($PSBoundParameters['Credential']) {
                    $CIMSessionParams.credential = $Credential
                }
				
                # Connectivity Test
                Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
                Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
                # CIM/WMI Connection
                #  WsMAN
                IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0') {
                    Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
                    $CimSession = New-CimSession @CIMSessionParams
                    $CimProtocol = $CimSession.protocol
                    Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
                }
				
                # DCOM
                ELSE {
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
			
            CATCH {
                Write-Warning -Message "PROCESS - Error on $Computer"
                $_.Exception.Message
                if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
                if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
                if ($ErrorProcessInvokeWmiMethod) { Write-Warning -Message "PROCESS Error - $ErrorProcessInvokeWmiMethod" }
            }#CATCH
        } # FOREACH
    }#PROCESS
    END {	
        if ($CimSession) {
            Write-Verbose -Message "END - Close CIM Session(s)"
            Remove-CimSession $CimSession
        }
        Write-Verbose -Message "END - Script is completed"
    }
}
## End Set-NetworkLevelAuthentication
## Begin Get-FolderSize
Function Get-FolderSize {

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

    [CmdletBinding(DefaultParameterSetName = 'FolderPath')]
    param 
    (
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'FolderPath')]
        [String[]]$FolderPath,
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = 'FolderPath')]
        [String]$FoldersOver,
        [Parameter(Mandatory = $false, Position = 2, ParameterSetName = 'FolderPath')]
        [switch]$Recurse

    )

    Begin {
        #$FoldersOver and $ZeroSizeFolders cannot be used together
        #Convert the size specified by Greaterhan parameter to Bytes
        $size = 1000000000 * $FoldersOver

    }
    ## End Get-FolderSize

    Process {
        #Check whether user has access to the folders.
	
	
        Try {
            Write-Host "Performing initial tasks, please wait... " -ForegroundColor Magenta
            $ColItems = If ($Recurse) { Get-ChildItem $FolderPath -Recurse -ErrorAction Stop } 
            Else { Get-ChildItem $FolderPath -ErrorAction Stop } 
		
        } 
        Catch [exception] {}
		
        #Calculate folder size
        If ($ColItems) {
            Write-Host "Calculating size of folders in $FolderPath. This may take sometime, please wait... " -ForegroundColor Magenta
            $Items = $ColItems | Where-Object { $_.PSIsContainer -eq $TRUE -and `
                @(Get-ChildItem -LiteralPath $_.Fullname -Recurse -ErrorAction SilentlyContinue | Where-Object { !$_.PSIsContainer }).Length -gt '0' }
        }
		

        ForEach ($i in $Items) {

            $subFolders = 
            If ($FoldersOver)
            { Get-ChildItem -Path $i.FullName -Recurse | Measure-Object -sum Length | Where-Object { $_.Sum -ge $size -and $_.Sum -gt 100000000 } }
            Else
            { Get-ChildItem -Path $i.FullName -Recurse | Measure-Object -sum Length | Where-Object { $_.Sum -gt 100000000 } } #added 25/12/2014: returns folders over 100MB
            #Return only values not equal to 0
            ForEach ($subFolder in $subFolders) {
                #If folder is less than or equal to 1GB, display in MB, If above 1GB, display in GB 
                $si = If (($subFolder.Sum -ge 1000000000)  ) { "{0:N2}" -f ($subFolder.Sum / 1GB) + " GB" } 
                ElseIf (($subFolder.Sum -lt 1000000000)  ) { "{0:N0}" -f ($subFolder.Sum / 1MB) + " MB" } 
                $Object = New-Object PSObject -Property @{            
                    'Folder Name' = $i.Name                
                    'Size'        = $si
                    'Full Path'   = $i.FullName          
                } 

                $Object | Select-Object 'Folder Name', 'Full Path', Size



            } 
            ## End Get-FolderSize

        }
        ## End Get-FolderSize


    }
    ## End Get-FolderSize
    End {

        Write-Host "Task completed...if nothing is displayed:
you may not have access to the path specified or 
all folders are less than 100 MB" -ForegroundColor Cyan


    }
    ## End Get-FolderSize

}
## End Get-FolderSize
## Begin Get-Software
Function Get-Software {

    [OutputType('System.Software.Inventory')]

    [Cmdletbinding()] 

    Param( 

        [Parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)] 

        [String[]]$Computername = $env:COMPUTERNAME

    )         

    Begin {

    }

    Process {     

        ForEach ($Computer in  $Computername) { 

            If (Test-Connection -ComputerName  $Computer -Count  1 -Quiet) {

                $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         

                ForEach ($Path in $Paths) { 

                    Write-Verbose  "Checking Path: $Path"

                    #  Create an instance of the Registry Object and open the HKLM base key 

                    Try { 

                        $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $Computer, 'Registry64') 

                    }
                    Catch { 

                        Write-Error $_ 

                        Continue 

                    } 

                    #  Drill down into the Uninstall key using the OpenSubKey Method 

                    Try {

                        $regkey = $reg.OpenSubKey($Path)  

                        # Retrieve an array of string that contain all the subkey names 

                        $subkeys = $regkey.GetSubKeyNames()      

                        # Open each Subkey and use GetValue Method to return the required  values for each 

                        ForEach ($key in $subkeys) {   

                            Write-Verbose "Key: $Key"

                            $thisKey = $Path + "\\" + $key 

                            Try {  

                                $thisSubKey = $reg.OpenSubKey($thisKey)   

                                # Prevent Objects with empty DisplayName 

                                $DisplayName = $thisSubKey.getValue("DisplayName")

                                If ($DisplayName -AND $DisplayName -notmatch '^Update  for|rollup|^Security Update|^Service Pack|^HotFix') {

                                    $Date = $thisSubKey.GetValue('InstallDate')

                                    If ($Date) {

                                        Try {

                                            $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)

                                        }
                                        Catch {

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

                                        $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32, 0)))

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

                                        Computername    = $Computer

                                        DisplayName     = $DisplayName

                                        Version         = $Version

                                        InstallDate     = $Date

                                        Publisher       = $Publisher

                                        UninstallString = $UninstallString

                                        InstallLocation = $InstallLocation

                                        InstallSource   = $InstallSource

                                        HelpLink        = $thisSubKey.GetValue('HelpLink')

                                        EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize') * 1024) / 1MB, 2))

                                    }

                                    $Object.pstypenames.insert(0, 'System.Software.Inventory')

                                    Write-Output $Object

                                }

                            }
                            Catch {

                                Write-Warning "$Key : $_"

                            }   

                        }

                    }
                    Catch {}   

                    $reg.Close() 

                }                  

            }
            Else {

                Write-Error  "$($Computer): unable to reach remote system!"

            }

        } 

    } 

}  
## End Get-Software
## Begin Get-AssetTagAndSerialNumber
Function Get-AssetTagAndSerialNumber {

    param  ( [string[]]$computerName = @('.') );

    $computerName | ForEach-Object {

        if ($_) {

            Get-WmiObject -ComputerName $_ Win32_SystemEnclosure | Select-Object __Server, SerialNumber, SMBiosAssetTag

        }

    }

}
## End Get-AssetTagAndSerialNumber
## Begin Clear-Memory
Function Clear-Memory {
    Get-Variable |
    Where-Object { $startupVariables -notcontains $_.Name } |
    ForEach-Object {
        try { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
        catch { }
    }
}
## End Clear-Memory
## Begin Remove-UserVariable
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
            catch { Write-Warning -Message "An error has occured. Error Details: $($_.Exception.Message)" }           
        }
    }
    else { Write-Warning -Message '$StartupVars has not been added to your PowerShell profile' }   
    $StartupVars = @()
    $StartupVars = Get-Variable | Select-Object -ExpandProperty Name
}
## End Remove-UserVariable
## Begin Enable-PSTranscriptionLogging
Function Enable-PSTranscriptionLogging {
    param(
        [Parameter(Mandatory)]
        [string]$OutputDirectory
    )

    # Registry path
    $basePath = 'HKLM:\SOFTWARE\WOW6432Node\Policies\Microsoft\Windows\PowerShell\Transcription'

    # Create the key if it does not exist
    if (-not (Test-Path $basePath)) {
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
## End Enable-PSTranscriptionLogging
## Begin Get-OlderFiles
Function Get-OlderFiles {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
 
    #Check and return if the provided Path not found
    if (-not (Test-Path -Path $Path) ) {
        Write-Error "Provided Path ($Path) not found"
        return
    }
    ## End Get-OlderFiles
 
    try {
        $files = Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $files) {
         
            #Skip directories as the current focus is only on files
            if ($file.PSIsContainer) {
                Continue
            }
 
            $last_modified = $file.Lastwritetime
            $time_diff_in_days = [math]::floor(((get-date) - $last_modified).TotalDays)
            $obj = New-Object -TypeName PSObject
            $obj | Add-Member -MemberType NoteProperty -Name FileName -Value $file.Name
            $obj | Add-Member -MemberType NoteProperty -Name FullPath -Value $file.FullName
            $obj | Add-Member -MemberType NoteProperty -Name AgeInDays -Value $time_diff_in_days
            $obj | Add-Member -MemberType NoteProperty -Name SizeInMB -Value $([Math]::Round(($file.Length / 1MB), 3))
            $obj
        }
    }
    catch {
        ## End Get-OlderFiles
        Write-Error "Error occurred. $_"
    }
}
## End Get-OlderFiles
## Begin Find-User
Function Find-User ($username) {
    $homeserver = ((get-aduser -id $username -prop homedirectory).Homedirectory -split "\\")[2]
    $query = "SELECT UserName,ComputerName,ActiveTime,IdleTime from win32_serversession WHERE UserName like '$username'"
    $results = Get-WmiObject -Namespace root\cimv2 -computer $homeServer -Query $query | Select-Object UserName, ComputerName, ActiveTime, IdleTime
    foreach ($result in $results) {
        $hostname = ""
        $hostname = [System.net.Dns]::GetHostEntry($result.ComputerName).hostname
        $result | Add-Member -Type NoteProperty -Name HostName -Value $hostname -force
        $result | Add-Member -Type NoteProperty -Name HomeServer -Value $homeServer -force
    }
    $results
}
## End Find-User
## Begin Invoke-WSUSCheckin
Function Invoke-WSUSCheckin($Computer) {
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
## End Invoke-WSUSCheckin
## Begin Get-EffectiveAccess
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

    begin {
        # requires -Modules ActiveDirectory
        $ErrorActionPreference = 'Stop'
        $GUIDMap = @{}
        $domain = Get-ADRootDSE
        $z = '00000000-0000-0000-0000-000000000000'
        $hash = @{
            SearchBase  = $domain.schemaNamingContext
            LDAPFilter  = '(schemaIDGUID=*)'
            Properties  = 'name', 'schemaIDGUID'
            ErrorAction = 'SilentlyContinue'
        }
        $schemaIDs = Get-ADObject @hash 

        $hash = @{
            SearchBase  = "CN=Extended-Rights,$($domain.configurationNamingContext)"
            LDAPFilter  = '(objectClass=controlAccessRight)'
            Properties  = 'name', 'rightsGUID'
            ErrorAction = 'SilentlyContinue'
        }
        $extendedRigths = Get-ADObject @hash

        foreach ($i in $schemaIDs) {
            if (-not $GUIDMap.ContainsKey([System.GUID]$i.schemaIDGUID)) {
                $GUIDMap.add([System.GUID]$i.schemaIDGUID, $i.name)
            }
        }
        foreach ($i in $extendedRigths) {
            if (-not $GUIDMap.ContainsKey([System.GUID]$i.rightsGUID)) {
                $GUIDMap.add([System.GUID]$i.rightsGUID, $i.name)
            }
        }
    }

    process {
        $result = [system.collections.generic.list[pscustomobject]]::new()
        $object = Get-ADObject $DistinguishedName
        $acls = (Get-ACL "AD:$object").Access
        
        foreach ($acl in $acls) {
            
            $objectType = if ($acl.ObjectType -eq $z) {
                'All Objects (Full Control)'
            }
            else {
                $GUIDMap[$acl.ObjectType]
            }

            $inheritedObjType = if ($acl.InheritedObjectType -eq $z) {
                'Applied to Any Inherited Object'
            }
            else {
                $GUIDMap[$acl.InheritedObjectType]
            }

            $result.Add(
                [PSCustomObject]@{
                    Name                  = $object.Name
                    IdentityReference     = $acl.IdentityReference
                    AccessControlType     = $acl.AccessControlType
                    ActiveDirectoryRights = $acl.ActiveDirectoryRights
                    ObjectType            = $objectType
                    InheritedObjectType   = $inheritedObjType
                    InheritanceType       = $acl.InheritanceType
                    IsInherited           = $acl.IsInherited
                })
        }
        
        if (-not $IncludeOrphan.IsPresent) {
            $result | Sort-Object IdentityReference |
            Where-Object { $_.IdentityReference -notmatch 'S-1-*' }
            return
        }

        return $result | Sort-Object IdentityReference
    }
}
## End Get-EffectiveAccess
## Begin Get-InstalledApplication
Function Get-InstalledApplication {
    [CmdletBinding()]
    Param(
        [Parameter(
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [String[]]$ComputerName = $ENV:COMPUTERNAME,

        [Parameter(Position = 1)]
        [String[]]$Properties,

        [Parameter(Position = 2)]
        [String]$IdentifyingNumber,

        [Parameter(Position = 3)]
        [String]$Name,

        [Parameter(Position = 4)]
        [String]$Publisher
    )
    Begin {
        Function IsCpuX86 ([Microsoft.Win32.RegistryKey]$hklmHive) {
            $regPath = 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
            $key = $hklmHive.OpenSubKey($regPath)

            $cpuArch = $key.GetValue('PROCESSOR_ARCHITECTURE')

            if ($cpuArch -eq 'x86') {
                return $true
            }
            else {
                return $false
            }
        }
    }
    Process {
        foreach ($computer in $computerName) {
            $regPath = @(
                'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
            )

            Try {
                $hive = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
                    [Microsoft.Win32.RegistryHive]::LocalMachine, 
                    $computer
                )
                if (!$hive) {
                    continue
                }
        
                # if CPU is x86 do not query for Wow6432Node
                if ($IsCpuX86) {
                    $regPath = $regPath[0]
                }

                foreach ($path in $regPath) {
                    $key = $hive.OpenSubKey($path)
                    if (!$key) {
                        continue
                    }
                    foreach ($subKey in $key.GetSubKeyNames()) {
                        $subKeyObj = $null
                        if ($PSBoundParameters.ContainsKey('IdentifyingNumber')) {
                            if ($subKey -ne $IdentifyingNumber -and 
                                $subkey.TrimStart('{').TrimEnd('}') -ne $IdentifyingNumber) {
                                continue
                            }
                        }
                        $subKeyObj = $key.OpenSubKey($subKey)
                        if (!$subKeyObj) {
                            continue
                        }
                        $outHash = New-Object -TypeName Collections.Hashtable
                        $appName = [String]::Empty
                        $appName = ($subKeyObj.GetValue('DisplayName'))
                        if ($PSBoundParameters.ContainsKey('Name')) {
                            if ($appName -notlike $name) {
                                continue
                            }
                        }
                        if ($appName) {
                            if ($PSBoundParameters.ContainsKey('Properties')) {
                                if ($Properties -eq '*') {
                                    foreach ($keyName in ($hive.OpenSubKey("$path\$subKey")).GetValueNames()) {
                                        Try {
                                            $value = $subKeyObj.GetValue($keyName)
                                            if ($value) {
                                                $outHash.$keyName = $value
                                            }
                                        }
                                        Catch {
                                            Write-Warning "Subkey: [$subkey]: $($_.Exception.Message)"
                                            continue
                                        }
                                    }
                                }
                                else {
                                    foreach ($prop in $Properties) {
                                        $outHash.$prop = ($hive.OpenSubKey("$path\$subKey")).GetValue($prop)
                                    }
                                }
                            }
                            $outHash.Name = $appName
                            $outHash.IdentifyingNumber = $subKey
                            $outHash.Publisher = $subKeyObj.GetValue('Publisher')
                            if ($PSBoundParameters.ContainsKey('Publisher')) {
                                if ($outHash.Publisher -notlike $Publisher) {
                                    continue
                                }
                            }
                            $outHash.ComputerName = $computer
                            $outHash.Path = $subKeyObj.ToString()
                            New-Object -TypeName PSObject -Property $outHash
                        }
                    }
                }
            }
            Catch {
                Write-Error $_
            }
        }
    }
    End {}
}
## End Get-InstalledApplication
## Begin Get-Uptime
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
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
 
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
                    ComputerName = $Computer
                    LastBoot     = $OS.ConvertToDateTime($OS.LastBootUpTime)
                    Uptime       = ([String]$Uptime.Days + " Days " + $Uptime.Hours + " Hours " + $Uptime.Minutes + " Minutes")
                }
 
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $Computer
                    LastBoot     = "Unable to Connect"
                    Uptime       = $_.Exception.Message.Split('.')[0]
                }
 
            }
            finally {
                $null = $OS
                $null = $Uptime
            }
        }
    }
 
    END {}
 
}
## End Get-Uptime
## Begin UnInstall-Modules
FFunction UnInstall-Modules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$RetentionMonths = 3
    )

    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Debug')) {
        $DebugPreference = 'Continue'
    }

    # Removed unused variable $CMDLetName

    # Get installed modules
    Write-Debug "Getting list of installed modules..."
    $Modules = Get-InstalledModule
    $Counter = 0

    foreach ($Module in $Modules) {
        Write-Debug "Found Module: $($Module.Name)"
    }

    foreach ($Module in $Modules) {
        Write-Host "`n"
        $ModuleVersions = Get-InstalledModule -Name $Module.Name -AllVersions
        $ModuleVersionsArray = @($ModuleVersions.Version)

        Write-Debug "Reviewing Module: $($Module.Name) - Versions Installed: $($ModuleVersionsArray.Count)"

        $VersionsToKeepArray = @()
        $MajorVersions = $ModuleVersionsArray.Major | Sort-Object -Unique
        $MinorVersions = $ModuleVersionsArray.Minor | Sort-Object -Unique

        foreach ($MajorVersion in $MajorVersions) {
            foreach ($MinorVersion in $MinorVersions) {
                $ReturnedVersion = Get-InstalledModule -Name $Module.Name -MaximumVersion "$MajorVersion.$MinorVersion.99999" -ErrorAction SilentlyContinue
                if ($ReturnedVersion) {
                    $VersionsToKeepArray += $ReturnedVersion.Version
                    $ModuleVersionsArray = $ModuleVersionsArray | Where-Object { $_ -ne $ReturnedVersion.Version }
                }
            }
        }

        # Remove older versions
        foreach ($Version in $ModuleVersionsArray) {
            Write-Debug "Removing Module: $($Module.Name) - Version: $Version"
            try {
                Uninstall-Module -Name $Module.Name -RequiredVersion $Version -ErrorAction Stop
                $Counter++
            } catch {
                Write-Warning "Failed to remove Module: $($Module.Name) - Version: $Version"
            }
        }

        # Evaluate removing previous builds older than retention period
        if ($VersionsToKeepArray.Count -gt 0) {
            $Oldest = ($VersionsToKeepArray | Measure-Object -Minimum).Minimum
            $Newest = ($VersionsToKeepArray | Measure-Object -Maximum).Maximum

            if ($Oldest -ne $Newest) {
                $ReturnedVersion = Get-InstalledModule -Name $Module.Name -RequiredVersion $Oldest -ErrorAction SilentlyContinue
                if ($ReturnedVersion -and ($ReturnedVersion.InstalledDate -lt (Get-Date).AddMonths(-$RetentionMonths))) {
                    try {
                        Uninstall-Module -Name $Module.Name -RequiredVersion $Oldest -ErrorAction Stop
                        $Counter++
                    } catch {
                        Write-Warning "Failed to remove oldest retained module version: $Oldest"
                    }
                } else {
                    Write-Debug "Module: $($Module.Name) - Version: $Oldest is within retention period, skipping removal."
                }
            }
        }
    }

    if ($Counter -gt 0) {
        Write-Debug "Removed $Counter module versions."
    }
}
## End UnInstall-Modules

## Begin Get-TaskPlus
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
        [switch]$Reverse,
        [VMware.VimAutomation.ViCore.Impl.V1.VIServerImpl[]]$Server = $global:DefaultVIServer,
        [switch]$Realtime,
        [switch]$Details,
        [switch]$Keys,
        [int]$WindowSize = 1000
    )
    begin {
        ## Begin Get-TaskDetails
        Function Get-TaskDetails {
            param(
                [VMware.Vim.TaskInfo[]]$Tasks
            )
            begin {
                $psV3 = $PSversionTable.PSVersion.Major -ge 3
            }
            ## End Get-TaskDetails
            process {
                $tasks | ForEach-Object {
                    if ($psV3) {
                        $object = [ordered]@{ }
                    }
                    ## End Get-TaskDetails
                    else {
                        $object = @{ }
                    }
                    ## End Get-TaskDetails
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
                    ## End Get-TaskDetails
                    if ($keys) {
                        $object.Add("Key", $_.Key)
                        $object.Add("ParentKey", $_.ParentTaskKey)
                        $object.Add("RootKey", $_.RootTaskKey)
                    }
                    ## End Get-TaskDetails
                    New-Object PSObject -Property $object
                }
                ## End Get-TaskDetails
            }
            ## End Get-TaskDetails
        }
        ## End Get-TaskDetails
        $filter = New-Object VMware.Vim.TaskFilterSpec
        if ($Alarm) {
            $filter.Alarm = $Alarm.ExtensionData.MoRef
        }
        ## End Get-TaskDetails
        if ($Entity) {
            $filter.Entity = New-Object VMware.Vim.TaskFilterSpecByEntity
            $filter.Entity.entity = $Entity.ExtensionData.MoRef
            if ($Recurse) {
                $filter.Entity.Recursion = [VMware.Vim.TaskFilterSpecRecursionOption]::all
            }
            ## End Get-TaskDetails
            else {
                $filter.Entity.Recursion = [VMware.Vim.TaskFilterSpecRecursionOption]::self
            }
            ## End Get-TaskDetails
        }
        ## End Get-TaskDetails
        if ($State) {
            $filter.State = $State
        }
        ## End Get-TaskDetails
        if ($Start -or $Finish) {
            $filter.Time = New-Object VMware.Vim.TaskFilterSpecByTime
            $filter.Time.beginTime = $Start
            $filter.Time.endTime = $Finish
            $filter.Time.timeType = [vmware.vim.taskfilterspectimeoption]::startedTime
        }
        ## End Get-TaskDetails
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
## End Get-TaskPlus
## Begin Update-Profile
Function Update-Profile {
    @(
        $Profile.AllUsersAllHosts,
        $Profile.AllUsersCurrentHost,
        $Profile.CurrentUserAllHosts,
        $Profile.CurrentUserCurrentHost
    ) | ForEach-Object {
        if (Test-Path $_) {
            Write-Verbose "Running $_"
            . $_
        }
    }    
}
## End Update-Profile
## Begin Remove-UserProfiles
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
        [parameter(mandatory = $true)]
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
        if ([string]::IsNullOrWhiteSpace($DeleteUsers)) {

            Write-Host "`nImproper value entered, excluding all users from deletion. You will need to re-run the command on $computer, if you wish to try again...`n"

        }

        #If Read-Host contains proper syntax (Starts with /id:) run command to delete specified user; DelProf will give a confirmation prompt
        elseif ($DeleteUsers -like "/id:*") {

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

    foreach ($computer in $computername) {
        if (Test-Connection -Quiet -Count 1 -Computer $Computer) { 

            UseDelProf2 
        }

        else {
            
            Write-Host "`nUnable to connect to $computer. Please try again..." -ForegroundColor Red
        }

    }
}
## End Remove-UserProfiles
## Begin Get-LastBoot
Function Get-LastBoot {
    <# 
  .SYNOPSIS 
  Retrieve last restart time for specified workstation(s) 

  .EXAMPLE 
  Get-LastBoot Computer123456 

  .EXAMPLE 
  Get-LastBoot 123456 
  #> 
    param([Parameter(Mandatory = $true)]
        [string[]] $ComputerName)
    if (($computername.length -eq 6)) {
        [int32] $dummy_output = $null;

        if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
            $computername = "Computer" + $computername.Replace("Computer", "")
        }	
    }

    $i = 0
    $j = 0

    foreach ($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

 
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{
            "Computer Name" = $Computer
            "Last Reboot"   = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}
## End Get-LastBoot
## Begin Get-LoggedOnUser
Function Get-LoggedOnUser {
    <# 
  .SYNOPSIS 
  Retrieve current user logged into specified workstations(s) 

  .EXAMPLE 
  Get-LoggedOnUser Computer123456 

  .EXAMPLE 
  Get-LoggedOnUser 123456 
  #> 
    Param([Parameter(Mandatory = $true)]
        [string[]] $ComputerName)
    if (($computername.length -eq 6)) {
        [int32] $dummy_output = $null;

        if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
            $computername = "Computer" + $computername.Replace("Computer", "")
        }	
    }
    write-host("")
    write-host("Gathering resources. Please wait...")
    write-host("")

    $i = 0
    $j = 0

    foreach ($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Logged On User..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

        $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer

        Write-Host "User Logged In: " $computerSystem.UserName "`n"
    }
}
## End Get-LoggedOnUser
## Begin CheckProcess
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
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = '')
    if (($computername.length -eq 6)) {
        [int32] $dummy_output = $null;

        if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
            $computername = "Computer" + $computername.Replace("Computer", "")
        }	
    }

    $Stamp = (Get-Date -Format G) + ":"
    $ComputerArray = @()

    ## Begin ChkProcess
    Function ChkProcess {

        $i = 0
        $j = 0

        foreach ($computer in $ComputerArray) {

            Write-Progress -Activity "Retrieving System Processes..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

            $getProcess = Get-Process -ComputerName $computer

            foreach ($Process in $getProcess) {
                
                [pscustomobject]@{
                    "Computer Name" = $computer
                    "Process Name"  = $Process.ProcessName
                    PID             = '{0:f0}' -f $Process.ID
                    Company         = $Process.Company
                    "CPU(s)"        = $Process.CPU
                    Description     = $Process.Description
                }           
            }
        } 


	
        foreach ($computer in $ComputerName) {	     
            If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
                $ComputerArray += $Computer
            }	
        }

        $chkProcess = ChkProcess | Sort-Object "Computer Name" | Select-Object "Computer Name", "Process Name", PID, Company, "CPU(s)", Description
        $DocPath = [environment]::getfolderpath("mydocuments") + "\Process-Report.csv"

        Switch ($CheckBox.IsChecked) {
            $true { $chkProcess | Export-Csv $DocPath -NoTypeInformation -Force; }
            default { $chkProcess | Out-GridView -Title "Processes"; }
        }

        if ($CheckBox.IsChecked -eq $true) {
            Try { 
                $listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
            } 

            Catch {
                #Do Nothing 
            }
        }
	
        else {
            Try {
                $listBox.Items.Add("$stamp Check Process output processed!`n")
            } 
            Catch {
                #Do Nothing 
            }
        }
    
    }
}
## End CheckProcess
## Begin Update-Sysinternals
Function Update-Sysinternals {
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
    [CmdletBinding()]
    param (
        # Path to the directory were sysinternals tools will be downloaded to 
        [Parameter(Mandatory = $true)]      
        [string]
        $Path 
    )
    
    begin {
        if (-not (Test-Path -Path $Path)) {
            Throw "The Path $_ does not exist"
        }
        else {
            $true
        }
        
        $uri = 'https://live.sysinternals.com/'
        $sysToolsPage = Invoke-WebRequest -Uri $uri
            
    }
    
    process {
        # create dir if it doesn't exist    
       
        Set-Location -Path $Path

        $sysTools = $sysToolsPage.Links.innerHTML | Where-Object -FilterScript { $_ -like "*.exe" -or $_ -like "*.chm" } 

        foreach ($sysTool in $sysTools) {
            Invoke-WebRequest -Uri "$uri/$sysTool" -OutFile $sysTool
        }
    } #process
}
## End Update-Sysinternals
## Begin Test-RegistryValue
Function Test-RegistryValue {

    param (

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]$Path,

        [parameter(Mandatory = $true)]
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
## End Test-RegistryValue
## Begin findfile($name)
Function findfile($name) {
    Get-ChildItem -Recurse -Filter "*${name}*" -ErrorAction SilentlyContinue | ForEach-Object {
        $place_path = $_.directory
        Write-Output "${place_path}\${_}"
    }
}
## End findfile($name)
## Begin sudo()
Function sudo() {
    Invoke-Elevated @args
}
## End sudo()
## Begin PSsed
Function PSsed($file, $find, $replace) {
	(Get-Content $file).replace("$find", $replace) | Set-Content $file
}
## End PSsed
## Begin PSsed-recursive
Function Set-ContentRecursive($filePattern, $find, $replace) {
    $files = Get-ChildItem . "$filePattern" -rec # -Exclude
    foreach ($file in $files) {
        (Get-Content $file.PSPath) |
        Foreach-Object { $_ -replace "$find", "$replace" } |
        Set-Content $file.PSPath
    }
}
## End PSsed-recursive
## Begin PSgrep
Function PSgrep {

    [CmdletBinding()]
    Param(
    
        # source file to grep
        [Parameter(Mandatory = $true)]
        [string]$SourceFileName, 

        # string to search for
        [Parameter(Mandatory = $true)]
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
            if ($_ -match $String) {
                $count ++
                if (!($OutputFile)) {
                    write-host $_
                }
                else {
                    $_ | Out-File -FilePath ".\$($OutputFile)" -Append -Force
                }

            }

        }

    }

    Write-Host "$($Count) matches found"
}
## End PSgrep
## Begin which($name)
Function which($name) {
    Get-Command $name | Select-Object -ExpandProperty Definition
}
## End which($name)
## Begin cut()
Function cut() {
    foreach ($part in $input) {
        $line = $part.ToString();
        $MaxLength = [System.Math]::Min(200, $line.Length)
        $line.subString(0, $MaxLength)
    }
}
## End cut()
## Begin Search-AllTextFiles
Function Search-AllTextFiles {
    param(
        [parameter(Mandatory = $true, position = 0)]$Pattern, 
        [switch]$CaseSensitive,
        [switch]$SimpleMatch
    );

    Get-ChildItem . * -Recurse -Exclude ('*.dll', '*.pdf', '*.pdb', '*.zip', '*.exe', '*.jpg', '*.gif', '*.png', '*.ico', '*.svg', '*.bmp', '*.psd', '*.cache', '*.doc', '*.docx', '*.xls', '*.xlsx', '*.dat', '*.mdf', '*.nupkg', '*.snk', '*.ttf', '*.eot', '*.woff', '*.tdf', '*.gen', '*.cfs', '*.map', '*.min.js', '*.data') | Select-String -Pattern:$pattern -SimpleMatch:$SimpleMatch -CaseSensitive:$CaseSensitive
}
## End Search-AllTextFiles
## Begin AddTo-7zip
Function AddTo-7zip($zipFileName) {
    BEGIN {
        #$7zip = "$($env:ProgramFiles)\7-zip\7z.exe"
        $7zip = Find-Program "\7-zip\7z.exe"
        if (!([System.IO.File]::Exists($7zip))) {
            throw "7zip not found";
        }
    }
    PROCESS {
        & $7zip a -tzip $zipFileName $_
    }
    END {
    }
}
## End AddTo-7zip
## Begin GoGo-PSExch
Function GoGo-PSExch {

    param(
        [Parameter( Mandatory = $false)]
        [string]$URL = "USONVSVREX01"
    )
    
    #$Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

    $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos #-Credential #$Credentials

    Import-PSSession $ExOPSession -AllowClobber

}
## End GoGo-PSExch

## Begin Get-FileAttribute
Function Get-FileAttribute {
    param($file, $attribute)
    $val = [System.IO.FileAttributes]$attribute;
    if ((Get-ChildItem $file -force).Attributes -band $val -eq $val) { $true; } else { $false; }
} 
## End Get-FileAttribute
## Begin Set-FileAttribute
Function Set-FileAttribute {
    param($file, $attribute)
    $file = (Get-ChildItem $file -force);
    $file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]$attribute).value__;
    if ($?) { $true; } else { $false; }
} 
## End Set-FileAttribute
## Begin LastBoot
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

        [Parameter(Mandatory = $true)]
        [String[]]$ComputerName,

        $i = 0,
        $j = 0
    )

    foreach ($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)
 
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{

            ComputerName = $Computer
            LastReboot   = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}#End LastBoot
## End LastBoot
## Begin SYSinfo
Function SYSinfo {
    <# 
.SYNOPSIS 
  Retrieve basic system information for specified workstation(s) 

.EXAMPLE 
  SYS Computer123456 
#> 

    param(

        [Parameter(Mandatory = $true)]
        [string[]] $ComputerName,
    
        $i = 0,
        $j = 0
    )

    $Stamp = (Get-Date -Format G) + ":"

    Function Systeminformation {
	
        foreach ($Computer in $ComputerName) {

            if (!([String]::IsNullOrWhiteSpace($Computer))) {

                if (Test-Connection -Quiet -Count 1 -Computer $Computer) {

                    Write-Progress -Activity "Getting Sytem Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

                    Start-Job -ScriptBlock { param($Computer) 

                        #Gather specified workstation information; CimInstance only works on 64-bit
                        $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer
                        $computerBIOS = Get-CimInstance CIM_BIOSElement -Computer $Computer
                        $computerOS = Get-CimInstance CIM_OperatingSystem -Computer $Computer
                        $computerCPU = Get-CimInstance CIM_Processor -Computer $Computer
                        $computerHDD = Get-CimInstance Win32_LogicalDisk -Computer $Computer -Filter "DeviceID = 'C:'"
    
                        [PSCustomObject] @{

                            ComputerName    = $computerSystem.Name
                            LastReboot      = $computerOS.LastBootUpTime
                            OperatingSystem = $computerOS.OSArchitecture + " " + $computerOS.caption
                            Model           = $computerSystem.Model
                            RAM             = "{0:N2}" -f [int]($computerSystem.TotalPhysicalMemory / 1GB) + "GB"
                            DiskCapacity    = "{0:N2}" -f ($computerHDD.Size / 1GB) + "GB"
                            TotalDiskSpace  = "{0:P2}" -f ($computerHDD.FreeSpace / $computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace / 1GB) + "GB)"
                            CurrentUser     = $computerSystem.UserName
                        }
                    } -ArgumentList $Computer
                }

                else {

                    Start-Job -ScriptBlock { param($Computer)  
                     
                        [PSCustomObject] @{

                            ComputerName    = $Computer
                            LastReboot      = "Unable to PING."
                            OperatingSystem = "$Null"
                            Model           = "$Null"
                            RAM             = "$Null"
                            DiskCapacity    = "$Null"
                            TotalDiskSpace  = "$Null"
                            CurrentUser     = "$Null"
                        }
                    } -ArgumentList $Computer                       
                }
            }

            else {
                 
                Start-Job -ScriptBlock { param($Computer)  
                     
                    [PSCustomObject] @{

                        ComputerName    = "Value is null."
                        LastReboot      = "$Null"
                        OperatingSystem = "$Null"
                        Model           = "$Null"
                        RAM             = "$Null"
                        DiskCapacity    = "$Null"
                        TotalDiskSpace  = "$Null"
                        CurrentUser     = "$Null"
                    }
                } -ArgumentList $Computer
            }
        } 
    }

    $SystemInformation = SystemInformation | Receive-Job -Wait | Select-Object ComputerName, CurrentUser, OperatingSystem, Model, RAM, DiskCapacity, TotalDiskSpace, LastReboot
    $DocPath = [environment]::getfolderpath("mydocuments") + "\SystemInformation-Report.csv"

    Switch ($CheckBox.IsChecked) {

        $true { 
            
            $SystemInformation | Export-Csv $DocPath -NoTypeInformation -Force 
        }

        default { 
            
            $SystemInformation | Out-GridView -Title "System Information"
        }
    }

    if ($CheckBox.IsChecked -eq $true) {

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
## End SYSinfo
## Begin NetMSG
Function NetMSG {
    <# 
.SYNOPSIS 
    Generate a pop-up window on specified workstation(s) with desired message 

.EXAMPLE 
    NetMSG Computer123456 
#> 
	
    param(

        [Parameter(Mandatory = $true)]
        [String[]] $ComputerName,

        [Parameter(Mandatory = $true, HelpMessage = 'Enter desired message')]
        [String]$MyMessage,

        [String]$User = [Environment]::UserName,

        [String]$UserJob = (Get-ADUser $User -Property Title).Title,
    
        [String]$CallBack = "$User | 5-2444 | $UserJob",

        $i = 0,
        $j = 0
    )

    Function SendMessage {

        foreach ($Computer in $ComputerName) {

            Write-Progress -Activity "Sending messages..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)         

            #Invoke local MSG command on specified workstation - will generate pop-up message for any user logged onto that workstation - *Also shows on Login screen, stays there for 100,000 seconds or until interacted with
            Invoke-Command -ComputerName $Computer { param($MyMessage, $CallBack, $User, $UserJob)
 
                MSG /time:100000 * /v "$MyMessage {$CallBack}"
            } -ArgumentList $MyMessage, $CallBack, $User, $UserJob -AsJob
        }
    }

    SendMessage | Wait-Job | Remove-Job

}#End NetMSG
## End NetMSG
## Begin InstallApplication
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

        [Parameter(Mandatory = $true, HelpMessage = "Enter Computername(s)")]
        [String[]]$Computername,

        [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter installer path(s)")]
        [String[]]$Path = $null,

        [Parameter(ValueFromPipeline = $true, HelpMessage = 'Enter remote destination: C$\Directory')]
        $Destination = "C$\TempApplications"
    )

    if ($null -eq $Path) {

        Add-Type -AssemblyName System.Windows.Forms

        $Dialog = New-Object System.Windows.Forms.OpenFileDialog
        $Dialog.InitialDirectory = "\\lasfs03\Software\Current Version\Deploy"
        $Dialog.Title = "Select Installation File(s)"
        $Dialog.Filter = "Installation Files (*.exe,*.msi,*.msp)|*.exe; *.msi; *.msp"        
        $Dialog.Multiselect = $true
        $Result = $Dialog.ShowDialog()

        if ($Result -eq 'OK') {

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
        foreach ($Computer in $Computername) {

            #If $Computer IS NOT null or only whitespace
            if (!([string]::IsNullOrWhiteSpace($Computer))) {

                #Test-Connection to $Computer
                if (Test-Connection -Quiet -Count 1 $Computer) {                                               
                     
                    #Create job on localhost
                    Start-Job { param($Computer, $Path, $Destination)

                        foreach ($P in $Path) {
                            
                            #Static Temp location
                            $TempDir = "\\$Computer\$Destination"

                            #Create $TempDir directory
                            if (!(Test-Path $TempDir)) {

                                New-Item -Type Directory $TempDir | Out-Null
                            }
                     
                            #Retrieve Leaf object from $Path
                            $FileName = (Split-Path -Path $P -Leaf)

                            #New Executable Path
                            $Executable = "C:\$(Split-Path -Path $Destination -Leaf)\$FileName"

                            #Copy needed installer files to remote machine
                            Copy-Item -Path $P -Destination $TempDir

                            #Install .EXE
                            if ($FileName -like "*.exe") {

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
                            elseif ($FileName -like "*.msi") {

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
                            elseif ($FileName -like "*.msp") { 
                                                                       
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
## End InstallApplication
## Begin Get-Icon
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
        [parameter(ValueFromPipelineByPropertyName = $True)]
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
            $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path) | 
            Add-Member -MemberType NoteProperty -Name FullName -Value $Path -PassThru
            If ($PSBoundParameters.ContainsKey('ToBytes')) {
                Write-Verbose "Retrieving bytes"
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $MemoryStream.ToArray()   
                $MemoryStream.Flush()  
                $MemoryStream.Dispose()           
            }
            ElseIf ($PSBoundParameters.ContainsKey('ToBitmap')) {
                $Icon.ToBitMap()
            }
            ElseIf ($PSBoundParameters.ContainsKey('ToBase64')) {
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $Bytes = $MemoryStream.ToArray()   
                $MemoryStream.Flush() 
                $MemoryStream.Dispose()
                [convert]::ToBase64String($Bytes)
            }
            Else {
                $Icon
            }
        }
        Else {
            Write-Warning "$Path does not exist!"
            Continue
        }
    }
}
## End Get-Icon
## Begin Get-MappedDrive
Function Get-MappedDrive {
    param (
        [string]$computername = "localhost"
    )
    Get-WmiObject -Class Win32_MappedLogicalDisk -ComputerName $computername | 
    Format-List DeviceId, VolumeName, SessionID, Size, FreeSpace, ProviderName
}
## End Get-MappedDrive
## Begin Get-UserLastLogonTime
Function Get-UserLastLogonTime {

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
    Begin {
        #List of users that are present on all PCs, these won't output unless specified to do so.
        $CommonUsers = ("NetworkService", "LocalService", "systemprofile")
    }
    
    #Process Pipeline
    Process {
        #ping the machine before trying to do anything
        if (Test-Connection $ComputerName -Count 2 -Quiet) {
            #try to get the OS version of the computer
            try { $OS = (gwmi  -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction stop).caption }
            catch {
                #had an error getting the WMI-object
                return New-Object psobject -Property @{
                    User        = "Error getting WMIObject Win32_OperatingSystem"
                    LastUseTime = get-date 0
                }
            }
            #make sure the OS retrieved is either Windows 7 or Windows 10 as this Function has not been set to work on other operating systems
            if ($OS.contains("Windows 10") -or $OS.Contains("Windows 7")) {
                try {
                    #try to get the WMiObject win32_UserProfile
                    $UserObjects = Get-WmiObject -ComputerName $ComputerName -Class Win32_userProfile -Property LocalPath, LastUseTime -ErrorAction Stop
                    $users = @()
                    #loop that handles all the data that came back from the WMI objects
                    forEach ($UserObject in $UserObjects) {
                        #extract the username from the local path, first find where the last slash is after, get everything after that into its own string
                        $i = $UserObject.localPath.LastIndexOf("\") + 1
                        $tempUserString = ""
                        while ($null -ne $UserObject.localPath.toCharArray()[$i]) {
                            $tempUserString += $UserObject.localPath.toCharArray()[$i]
                            $i++
                        }
                        #if list common users is turned on, skip this block
                        if (!$listCommonUsers) {
                            #check if the user extracted from the local path is a common user
                            $isCommonUser = $false
                            forEach ($userName in $CommonUsers) { 
                                if ($userName -eq $tempUserString) {
                                    $isCommonUser = $true
                                    break
                                }
                            }
                        }
                        #if the user is one of the users specified in the common users, skip it unless otherwise specified
                        if ($isCommonUser) { continue }
                        #check to see if the user has a timestamp for there last logon 
                        if ($null -ne $UserObject.LastUseTime) {
                            #This converts the string timestamp to a DateTime object
                            $TempUserLastUseTime = ([WMI] '').ConvertToDateTime($userObject.LastUseTime)
                        }
                        #the user had a local path but no timestamp was found for last logon
                        else { $TempUserLastUseTime = Get-Date 0 }
                        #add user to array
                        $users += New-Object -TypeName psobject -Property @{
                            User        = $tempUserString
                            LastUseTime = $TempUserLastUseTime
                        }
                    }
                }
                catch {
                    #error trying to retrieve WMI obeject win32_userProfile
                    return New-Object psobject -Property @{
                        User        = "Error getting WMIObject Win32_userProfile"
                        LastUseTime = get-date 0
                    }
                }
            }
            else {
                #OS version was not compatible
                return New-Object psobject -Property @{
                    User        = "Operating system $OS is not compatible with this Function."
                    LastUseTime = get-date 0
                }
            }
        }
        else {
            #Computer was not pingable
            return New-Object psobject -Property @{
                User        = "Can't Ping"
                LastUseTime = get-date 0
            }
        }

        #check to see if any users came out of the main Function
        if ($users.count -eq 0) {
            $users += New-Object -TypeName psobject -Property @{
                User        = "NoLoggedOnUsers"
                LastUseTime = get-date 0
            }
        }
        #sort the user array by the last time they logged in
        else { $users = $users | Sort-Object -Property LastUseTime -Descending }
        #main output block
        #if List all users was chosen, output the full list of users found
        if ($ListAllUsers) { return $users }
        #if get last user was chosen, output the last user to log on the computer
        elseif ($GetLastUser) { return ($users[0]) }
        else {
            #see if the user specified ever logged on
            ForEach ($Username in $users) {
                if ($Username.User -eq $user) { return ($Username) }            
            }

            #user did not log on
            return New-Object psobject -Property @{
                User        = "$user"
                LastUseTime = get-date 0
            }
        }
    }
    #End Pipeline
    End { Write-Verbose "Function get-UserLastLogonTime is complete" }
}
## End Get-UserLastLogonTime
## Begin Remove-AliasFromScript
Function Remove-AliasFromScript {
    Get-Alias | 
    Select-Object Name, Definition | 
    ForEach-Object -Begin { $a = @{} } -Process { $a.Add($_.Name, $_.Definition) } -End {}

    $b = $errors = $null
    $b = $psISE.CurrentFile.Editor.Text

    [system.management.automation.psparser]::Tokenize($b, [ref]$errors) |
    Where-Object { $_.Type -eq "command" } |
    ForEach-Object `
    {
        if ($a.($_.Content)) {
            $b = $b -replace
            ('(?<=(\W|\b|^))' + [regex]::Escape($_.content) + '(?=(\W|\b|$))'),
            $a.($_.content)
        }
    }

    $ScriptWithoutAliases = $psISE.CurrentPowerShellTab.Files.Add()
    $ScriptWithoutAliases.Editor.Text = $b
    $ScriptWithoutAliases.Editor.SetCaretPosition(1, 1)
    $ScriptWithoutAliases.Editor.EnsureVisible(1)  
}
## End Remove-AliasFromScript
## Begin Replace-SpacesWithTabs
Function Replace-SpacesWithTabs {
    param
    (
        [int]$spaces = 2
    ) 
  
    $tab = "`t"
    $space = " " * $spaces
    $text = $psISE.CurrentFile.Editor.Text

    $newText = ""
  
    foreach ($line in $text -split [Environment]::NewLine) {
        if ($line -match "\S") {
            $pos = $line.IndexOf($Matches[0])
            $indentation = $line.SubString(0, $pos)
            $remainder = $line.SubString($pos)
      
            $replaced = $indentation -replace $space, $tab
      
            $newText += $replaced + $remainder + [Environment]::NewLine
        }
        else {
            $newText += $line + [Environment]::NewLine
        }

        $psISE.CurrentFile.Editor.Text = $newText
    }
}
## End Replace-SpacesWithTabs
## Begin Replace-TabsWithSpaces
Function Replace-TabsWithSpaces {
    param
    (
        [int]$spaces = 2
    )   
  
    $tab = "`t"
    $space = " " * $spaces
    $text = $psISE.CurrentFile.Editor.Text

    $newText = ""
  
    foreach ($line in $text -split [Environment]::NewLine) {
        if ($line -match "\S") {
            $pos = $line.IndexOf($Matches[0])
            $indentation = $line.SubString(0, $pos)
            $remainder = $line.SubString($pos)
      
            $replaced = $indentation -replace $tab, $space
      
            $newText += $replaced + $remainder + [Environment]::NewLine
        }
        else {
            $newText += $line + [Environment]::NewLine
        }

        $psISE.CurrentFile.Editor.Text = $newText
    }
}
## End Replace-TabsWithSpaces
## Begin Indent-SelectedText
Function Indent-SelectedText {
    param
    (
        [int]$spaces = 2
    )
  
    $tab = " " * $space
    $text = $psISE.CurrentFile.Editor.SelectedText

    $newText = ""
  
    foreach ($line in $text -split [Environment]::NewLine) {
        $newText += $tab + $line + [Environment]::NewLine
    }

    $psISE.CurrentFile.Editor.InsertText($newText)
}
## End Indent-SelectedText
## Begin Add-RemarkedText
Function Add-RemarkedText {
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

    foreach ($l in $text -Split [Environment]::NewLine) {
        $newText += "{0}{1}" -f ("#" + $l), [Environment]::NewLine
    }

    $psISE.CurrentFile.Editor.InsertText($newText)
}
## End Add-RemarkedText
## Begin Remove-RemarkedText
Function Remove-RemarkedText {
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

    foreach ($l in $text -Split [Environment]::NewLine) {
        $newText += "{0}{1}" -f ($l -Replace '#', ''), [Environment]::NewLine
    }

    $psISE.CurrentFile.Editor.InsertText($newText)
}
## End Remove-RemarkedText
## Begin AbortScript
Function AbortScript {
    $Word.Quit()
    Write-Verbose "$(Get-Date): System Cleanup"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject( $Word ) | Out-Null
    If ( Get-Variable -Name Word -Scope Global ) {
        Remove-Variable -Name word -Scope Global
    }
    [GC]::Collect() 
    [GC]::WaitForPendingFinalizers()
    Write-Verbose "$(Get-Date): Script has been aborted"
    $ErrorActionPreference = $SaveEAPreference
    Exit
}
## End AbortScript
## Begin Add-OSCPicture
Function Add-OSCPicture {
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
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias('wordpath')]
        [String]$WordDocumentPath,
        [Parameter(Mandatory = $true, Position = 1)]
        [Alias('imgpath')]
        [String]$ImageFolderPath
    )

    If (Test-Path -Path $WordDocumentPath) {
        If (Test-Path -Path $ImageFolderPath) {
            $WordExtension = (Get-Item -Path $WordDocumentPath).Extension
            If ($WordExtension -like ".doc" -or $WordExtension -like ".docx") {
                $ImageFiles = Get-ChildItem -Path $ImageFolderPath -Recurse -Include *.emf, *.wmf, *.jpg, *.jpeg, *.jfif, *.png, *.jpe, *.bmp, *.dib, *.rle, *.gif, *.emz, *.wmz, *.pcz, *.tif, *.tiff, *.eps, *.pct, *.pict, *.wpg

                If ($ImageFiles) {
                    #Create the Word application object
                    $WordAPP = New-Object -ComObject Word.Application
                    $WordDoc = $WordAPP.Documents.Open("$WordDocumentPath")

                    Foreach ($ImageFile in $ImageFiles) {
                        $ImageFilePath = $ImageFile.FullName

                        $Properties = @{'ImageName' = $ImageFile.Name
                            'Action(Insert)'        = Try {
                                $WordAPP.Selection.EndKey(6) | Out-Null
                                $WordApp.Selection.InlineShapes.AddPicture("$ImageFilePath") | Out-Null
                                $WordApp.Selection.InsertNewPage() #insert new page to word
                                "Finished"
                            }
                            Catch {
                                "Unfinished"
                            }
                        }

                        $objWord = New-Object -TypeName PSObject -Property $Properties
                        $objWord
                    }

                    $WordDoc.Save()
                    $WordDoc.Close()
                    $WordAPP.Quit()#release the object
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordAPP) | Out-Null
                    Remove-Variable WordAPP
                }
                Else {
                    Write-Warning "There is no image in this '$ImageFolderPath' folder."
                }
            }
            Else {
                Write-Warning "There is no word document file in this '$WordDocumentPath' folder."
            }
        }
        Else {
            Write-Warning "Cannot find path '$ImageFolderPath' because it does not exist."
        }
    }
    Else {
        Write-Warning "Cannot find path '$WordDocumentPath' because it does not exist."
    }
}
## End Add-OSCPicture
## Begin Connect-Office365
Function Connect-Office365 {
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
    BEGIN {
        TRY {
            #Modules
            IF (-not (Get-Module -Name MSOnline -ListAvailable)) {
                Write-Verbose -Message "BEGIN - Import module Azure Active Directory"
                Import-Module -Name MSOnline -ErrorAction Stop -ErrorVariable ErrorBeginIpmoMSOnline
            }
			
            IF (-not (Get-Module -Name LyncOnlineConnector -ListAvailable)) {
                Write-Verbose -Message "BEGIN - Import module Lync Online"
                Import-Module -Name LyncOnlineConnector -ErrorAction Stop -ErrorVariable ErrorBeginIpmoLyncOnline
            }
        }
        CATCH {
            Write-Warning -Message "BEGIN - Something went wrong!"
            IF ($ErrorBeginIpmoMSOnline) {
                Write-Warning -Message "BEGIN - Error while importing MSOnline module"
            }
            IF ($ErrorBeginIpmoLyncOnline) {
                Write-Warning -Message "BEGIN - Error while importing LyncOnlineConnector module"
            }
			
            Write-Warning -Message $error[0].exception.message
        }
    }
    PROCESS {
        TRY {
			
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
            Import-PSSession -Session $O365PS �Prefix ExchCloud
			
            # LYNC ONLINE (LyncOnlineConnector)
            Write-Verbose -Message "PROCESS - Create session to Lync online"
            $lyncsession = New-CsOnlineSession �Credential $O365cred -ErrorAction Stop -ErrorVariable ErrorConnectExchange
            Import-PSSession -Session $lyncsession -Prefix LyncCloud
			
            # SHAREPOINT ONLINE
            #Connect-SPOService -Url https://contoso-admin.sharepoint.com �credential $O365cred
        }
        CATCH {
            Write-Warning -Message "PROCESS - Something went wrong!"
            IF ($ErrorCredential) {
                Write-Warning -Message "PROCESS - Error while gathering credential"
            }
            IF ($ErrorConnectMSOL) {
                Write-Warning -Message "PROCESS - Error while connecting to Azure AD"
            }
            IF ($ErrorConnectExchange) {
                Write-Warning -Message "PROCESS - Error while connecting to Exchange Online"
            }
            IF ($ErrorConnectLync) {
                Write-Warning -Message "PROCESS - Error while connecting to Lync Online"
            }
			
            Write-Warning -Message $error[0].exception.message
        }
    }
}
## End Connect-Office365
## Begin New-CimSmartSession
Function New-CimSmartSession {
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
		
        [System.Management.Automation.PSCredential]
        $Credential
    )
	
    BEGIN {
        # Default Verbose/Debug message
        Function Get-DefaultMessage {
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
	
    PROCESS {
        FOREACH ($Computer in $ComputerName) {
            Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Test-Connection")
            IF (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                $CIMSessionSplatting.ComputerName = $Computer
				
				
                # WSMAN Protocol
                IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: ([3-9]|[1-9][0-9]+)\.[0-9]+') {
                    TRY {
                        #WSMAN (Default when using New-CimSession)
                        Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Connecting using WSMAN protocol (Default, requires at least PowerShell v3.0)")
                        New-CimSession @CIMSessionSplatting -errorVariable ErrorProcessNewCimSessionWSMAN
                    }
                    CATCH {
                        IF ($ErrorProcessNewCimSessionWSMAN) { Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Can't Connect using WSMAN protocol") }
                        Write-Warning -Message (Get-DefaultMessage -Message $Error.Exception.Message)
                    }
                }
				
                ELSE {
                    # DCOM Protocol
                    $CIMSessionSplatting.SessionOption = $CIMSessionOption
					
                    TRY {
                        Write-Verbose -Message (Get-DefaultMessage -Message "$Computer - Connecting using DCOM protocol")
                        New-CimSession @SessionParams -errorVariable ErrorProcessNewCimSessionDCOM
                    }
                    CATCH {
                        IF ($ErrorProcessNewCimSessionDCOM) { Write-Warning -Message (Get-DefaultMessage -Message "$Computer - Can't connect using DCOM protocol either") }
                        Write-Warning -Message (Get-DefaultMessage -Message $Error.Exception.Message)
                    }
                    FINALLY {
                        # Remove the CimSessionOption for the DCOM protocol for the next computer
                        $CIMSessionSplatting.Remove('CIMSessionOption')
                    }
                }#ELSE
            }#Test-Connection
        }#FOREACH
    }#PROCESS
}#Function
## End New-CimSmartSession
## Begin New-DjoinFile
Function New-DjoinFile {
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
        [System.IO.FileInfo]$DestinationFile
    PARAM (
        [Parameter(Mandatory = $true)]
        [System.String]$Blob,
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$DestinationFile = "c:\temp\djoin.tmp"
    )
	
    PROCESS {
        TRY {
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
        CATCH {
            $Error[0]
        }
    }
}
## End New-DjoinFile
## Begin New-IsoFile
Function New-IsoFile {  
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
  
    [CmdletBinding(DefaultParameterSetName = 'Source')]Param( 
        [parameter(Position = 1, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Source')]$Source,  
        [parameter(Position = 2)][string]$Path = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",  
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })][string]$BootFile = $null, 
        [ValidateSet('CDR', 'CDRW', 'DVDRAM', 'DVDPLUSR', 'DVDPLUSRW', 'DVDPLUSR_DUALLAYER', 'DVDDASHR', 'DVDDASHRW', 'DVDDASHR_DUALLAYER', 'DISK', 'DVDPLUSRW_DUALLAYER', 'BDR', 'BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER', 
        [string]$Title = (Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
        [switch]$Force, 
        [parameter(ParameterSetName = 'Clipboard')][switch]$FromClipboard 
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
## End New-IsoFile
'@  
        } 
  
        if ($BootFile) { 
            if ('BDR', 'BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
      ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type = 1 }).Open()  # adFileTypeBinary 
            $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
      ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
        } 
 
        $MediaType = @('UNKNOWN', 'CDROM', 'CDR', 'CDRW', 'DVDROM', 'DVDRAM', 'DVDPLUSR', 'DVDPLUSRW', 'DVDPLUSR_DUALLAYER', 'DVDDASHR', 'DVDDASHRW', 'DVDDASHR_DUALLAYER', 'DISK', 'DVDPLUSRW_DUALLAYER', 'HDDVDROM', 'HDDVDR', 'HDDVDRAM', 'BDROM', 'BDR', 'BDRE') 
 
        Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))" 
    ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName = $Title }).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 
  
        if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break } 
    }  
 
    Process { 
        if ($FromClipboard) { 
            if ($PSVersionTable.PSVersion.Major -lt 5) { Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break } 
            $Source = Get-Clipboard -Format FileDropList 
        } 
 
        foreach ($item in $Source) { 
            if ($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) { 
                $item = Get-Item -LiteralPath $item 
            } 
 
            if ($item) { 
                Write-Verbose -Message "Adding item to the target image: $($item.FullName)" 
                try { $Image.Root.AddTree($item.FullName, $true) } catch { Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') } 
            } 
        } 
    } 
 
    End {  
        if ($Boot) { $Image.BootImageOptions = $Boot }  
        $Result = $Image.CreateResultImage()  
        [ISOFile]::Create($Target.FullName, $Result.ImageStream, $Result.BlockSize, $Result.TotalBlocks) 
        Write-Verbose -Message "Target image ($($Target.FullName)) has been created" 
        $Target 
    } 
} 
## End New-IsoFile
## Begin New-Password
Function New-Password {
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
        [ValidateRange(1, 256)]
        [Int]$Count = 1
    )#PARAM
	
    BEGIN {
        # Create ScriptBlock with the ASCII Char Codes
        $PasswordCharCodes = { 33..126 }.invoke()
		
		
        # Exclude some ASCII Char Codes from the ScriptBlock
        #  Excluded characters are ",',.,/,1,<,>,`,O,0,l,|
        #  See http://www.asciitable.com/ for mapping
        34, 39, 46, 47, 49, 60, 62, 96, 48, 79, 108, 124 | ForEach-Object { [void]$PasswordCharCodes.Remove($_) }
        $PasswordChars = [char[]]$PasswordCharCodes
    }#BEGIN

    PROCESS {
        1..$count | ForEach-Object {
            # Password of 4 characters or longer
            IF ($Length -gt 4) {
			
                DO {
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
            ELSE {
                $NewPassWord = $(foreach ($i in 1..$length) { Get-Random -InputObject $PassWordChars }) -join ''
            }#ELSE
		
            # Output a new password
            Write-Output $NewPassword
        }
    } #PROCESS
    END {
        # Cleanup
        Remove-Variable -Name NewPassWord -ErrorAction 'SilentlyContinue'
    } #END
} #Function
## End New-Password
## Begin New-RandomPassword
Function New-RandomPassword {
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
	
    BEGIN {
        Add-Type -AssemblyName System.web;
    }
	
    PROCESS {
        1..$Count | ForEach-Object {
            [System.Web.Security.Membership]::GeneratePassword($Length, $NumberOfNonAlphanumericCharacters)
        }
    }
}
## End New-RandomPassword
## Begin New-ScriptMessage
Function New-ScriptMessage {
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
	
    PROCESS {
        $DateFormat = Get-Date -Format $DateFormat
        $MyCommand = (Get-Variable -Scope $FunctionScope -Name MyInvocation -ValueOnly).MyCommand.Name
        IF ($MyCommand) {
            $String = "[$DateFormat][$MyCommand]"
        } #IF
        ELSE {
            $String = "[$DateFormat]"
        } #Else
		
        IF ($PSBoundParameters['Block']) {
            $String += "[$Block]"
        }
        Write-Output "$String $Message"
    } #Process
}
## End New-ScriptMessage
## Begin New-Shortcut
Function New-Shortcut {
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
            position = 0,
            mandatory = 1,
            ValueFromPipeline = 1,
            ValueFromPipeLineByPropertyName = 1)]
        [validateScript({ $_ | % { Test-Path $_ } })]
        [string[]]
        $TargetPaths,

        # set shortcut Directory to create shortcut. Default is user Desktop.
        [parameter(
            position = 1,
            mandatory = 0,
            ValueFromPipeLineByPropertyName = 1)]
        [validateScript({ -not(Test-Path $_) })]
        [string]
        $ShortcutDirectory = "$env:USERPROFILE\Desktop",

        # Set Description for shortcut.
        [parameter(
            position = 2,
            mandatory = 0,
            ValueFromPipeLineByPropertyName = 1)]
        [string]
        $Description,

        # set if you want to show create shortcut result
        [parameter(
            position = 3,
            mandatory = 0)]
        [switch]
        $PassThru
    )

    begin {
        $extension = ".lnk"
        $wsh = New-Object -ComObject Wscript.Shell
    }

    process {
        foreach ($TargetPath in $TargetPaths) {
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

            if ($PSBoundParameters.PassThru) {
                Write-Verbose ("Show Result for shortcut result for target file name '{0}'" -f $TargetPath)
                $shortCut
            }
        }
    }

    end {
    }
}
## End New-Shortcut
## Begin Query-SoftwareInstalled
Function Query-SoftwareInstalled {
    [CmdletBinding (SupportsShouldProcess = $True)]
    Param
    (
        [Parameter (Mandatory = $True,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            HelpMessage = "Input a domain OU structure to query for software installed.`r`nE.g.- `"OU=1stOUName,OU=2ndOUName,DC=LowLevelDomain,DC=MidLevelDomain,DC=TopLevelDomain`"")]
        [Alias('OU')]
        [string]$OUStructure,

        [Parameter (Mandatory = $True,
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False,
            HelpMessage = "Input the software name that is to be queried.`r`n(Be sure it matches the software's name listed in the registry.)")]
        [Alias('Install')]
        [string[]]$Software,

        [Parameter (Mandatory = $False,
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False,
            HelpMessage = "Input the number of days before a computer account is considered 'Inactive'.`r`nE.g.- `"30`"")]
        [Alias('Days')]
        [int32]$InactivityThreshold = "30",

        [Parameter (Mandatory = $False,
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False,
            HelpMessage = "Input a directory to output the CSV file.")]
        [Alias('Folder')]
        [string]$OutputPath = "$env:USERPROFILE\Desktop\$Software-Machines"
    )
    $date = Get-Date -Format MMM-dd-yyyy
    $time = [DateTime]::Now
    $Computers = Get-ADComputer -Filter { LastLogonTimeStamp -lt $time } -SearchBase "$OUStructure" -Properties 'Name', 'OperatingSystem', 'CanonicalName', 'LastLogonTimeStamp'
    If ($OutputPath) {
        If (-not (Test-Path -Path $OutputPath)) {
            New-Item -Path "$OutputPath" -ItemType Directory
        }
    }

    ForEach ($Computer in $Computers) {
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
        If ($LogonTime -gt (Get-Date).AddDays( - ($InactivityThreshold))) {
            $Name = $($Computer.Name)
            $NetConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $($Computer.Name) -ErrorAction SilentlyContinue -ErrorVariable NetAdapterError | Where { $_.IPEnabled -eq $true }
            If ($NetAdapterError -like "*The RPC server is unavailable*") {
                $IPAddress = "ERROR: Remote connection to $($Computer.Name)`'s network adapter failed."
            }
            Else {
                ForEach ($AdapterItem in $NetConfig) {
                    $MAC = $AdapterItem.MACAddress
                    $IPAddress = $AdapterItem.IPAddress | Where { $_ -like "172.*" }
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
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $($Computer.Name))
            $ErrorActionPreference = $PrevErrorActionPreference
            If ($error[0] -like "*Exception calling `"OpenRemoteBaseKey`"*" -and $error[0] -like "*`"The network path was not found.*") {
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
            Else {
                #Drill down into the Uninstall key using the OpenSubKey Method
                $regkey = $reg.OpenSubKey($UninstallKey)
                #Retrieve an array of strings that contain all the subkey names
                $subkeys = $regkey.GetSubKeyNames() 
                #Open each Subkey and use GetValue Method to return the required values for each
                ForEach ($key in $subkeys) {
                    $thisKey = $UninstallKey + "\\" + $key
                    $thisSubKey = $reg.OpenSubKey($thisKey)
                    If ($($thisSubKey.GetValue("DisplayName")) -like "*$Software*") {
                        $obj = New-Object PSObject
                        $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
                        $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
                        $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
                        $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
                        $subkeyarray += $obj
                        $obj = $null
                    }
                }
                If ($subkeyarray.DisplayName -notlike "*No Network connection*" -and $subkeyarray.DisplayName -notlike "*$Software*") {
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
        Else {
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
## End Query-SoftwareInstalled
## Begin Remove-HashTableEmptyValue
Function Remove-HashTableEmptyValue {
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
        if ($HashTable[$_] -eq "" -or $HashTable[$_] -eq $null) {
            Write-Verbose -Message "[Remove-HashTableEmptyValue][PROCESS] - Property: $_ removing..."
            [void]$HashTable.Remove($_)
            Write-Verbose -Message "[Remove-HashTableEmptyValue][PROCESS] - Property: $_ removed"
        }
    }
}
## End Remove-HashTableEmptyValue
## Begin Remove-PSObjectEmptyOrNullProperty
Function Remove-PSObjectEmptyOrNullProperty {
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
    PROCESS {
        $PsObject.psobject.Properties |
        Where-Object { -not $_.value } |
        ForEach-Object {
            $PsObject.psobject.Properties.Remove($_.name)
        }
    }
}
## End Remove-PSObjectEmptyOrNullProperty
## Begin Remove-PSObjectProperty
Function Remove-PSObjectProperty {
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
    PROCESS {
        Foreach ($item in $Property) {
            $PSObject.psobject.Properties.Remove("$item")
        }
    }
}
## End Remove-PSObjectProperty
## Begin Remove-StringDiacritic
Function Remove-StringDiacritic {
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
	PS C:\> Remove-StringDiacritic "L'�t� de Rapha�l"
	
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
	
    FOREACH ($StringValue in $String) {
        Write-Verbose -Message "$StringValue"
        try {	
            # Normalize the String
            $Normalized = $StringValue.Normalize($NormalizationForm)
            $NewString = New-Object -TypeName System.Text.StringBuilder
			
            # Convert the String to CharArray
            $normalized.ToCharArray() |
            ForEach-Object -Process {
                if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
                    [void]$NewString.Append($psitem)
                }
            }

            #Combine the new string chars
            Write-Output $($NewString -as [string])
        }
        Catch {
            Write-Error -Message $Error[0].Exception.Message
        }
    }
}
## End Remove-StringDiacritic
## Begin Remove-StringLatinCharacter
Function Remove-StringLatinCharacter {
    <#
.SYNOPSIS
    Function to remove diacritics from a string
.PARAMETER String
	Specifies the String that will be processed
.EXAMPLE
    Remove-StringLatinCharacter -String "L'�t� de Rapha�l"

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
        [Parameter(ValueFromPipeline = $true)]
        [System.String[]]$String
    )
    PROCESS {
        FOREACH ($StringValue in $String) {
            Write-Verbose -Message "$StringValue"

            TRY {
                [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($StringValue))
            }
            CATCH {
                Write-Error -Message $Error[0].exception.message
            }
        }
    }
}
## End Remove-StringLatinCharacter
## Begin Remove-StringSpecialCharacter
Function Remove-StringSpecialCharacter {
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
    PROCESS {
        IF ($PSBoundParameters["SpecialCharacterToKeep"]) {
            $Regex = "[^\p{L}\p{Nd}"
            Foreach ($Character in $SpecialCharacterToKeep) {
                $Regex += "/$character"
            }
			
            $Regex += "]+"
        } #IF($PSBoundParameters["SpecialCharacterToKeep"])
        ELSE { $Regex = "[^\p{L}\p{Nd}]+" }
		
        FOREACH ($Str in $string) {
            Write-Verbose -Message "Original String: $Str"
            $Str -replace $regex, ""
        }
    } #PROCESS
}
## End Remove-StringSpecialCharacter
## Begin Remove-UserProfile
Function Remove-UserProfile {
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
        [Parameter(Position = 0, Mandatory = $false)]
        [String[]]$Exclude,
        [Parameter(Position = 1, Mandatory = $false)]
        [DateTime]$Before,
        [Parameter(Position = 2, Mandatory = $false)]
        [Switch]$DirectoryCleanup
    )

    Write-Verbose "Gathering List of Profiles on $env:COMPUTERNAME to Remove..."

    $userProfileFilter = "Loaded = 'False' AND Special = 'False'"
    $cleanupExclusions = @("Public", "Default")

    if ($Exclude) {
        foreach ($exclusion in $Exclude) {
            $userProfileFilter += "AND NOT LocalPath LIKE '%$exclusion'"
            $cleanupExclusions += $exclusion
        }
    }

    if ($Before) {
        $userProfileFilter += "AND LastUseTime < '$Before'"

        $keepUserProfileFilter = "Special = 'False' AND LastUseTime >= '$Before'"
        $profilesToKeep = Get-WmiObject -Class Win32_UserProfile -Filter $keepUserProfileFilter -ErrorAction Stop

        foreach ($profileToKeep in $profilesToKeep) {
            try {
                $userSID = New-Object -TypeName System.Security.Principal.SecurityIdentifier($($profileToKeep.SID))
                $userName = $userSID.Translate([System.Security.Principal.NTAccount])
                
                $keepUserName = $userName.Value -replace ".*\\", ""
                $cleanupExclusions += $keepUserName
            }
            catch [System.Security.Principal.IdentityNotMappedException] {
                Write-Warning "Cannot Translate SID to UserName - Not Adding Value to Exceptions List"
            }
        }
    }

    $profilesToDelete = Get-WmiObject -Class Win32_UserProfile -Filter $userProfileFilter -ErrorAction Stop

    if ($DirectoryCleanup) {
        $usersChildItem = Get-ChildItem -Path "C:\Users" -Exclude $cleanupExclusions

        foreach ($usersChild in $usersChildItem) {
            if ($profilesToDelete.LocalPath -notcontains $usersChild.FullName) {    
                try {
                    Write-Verbose "Additional Directory Cleanup - Removing $($usersChild.Name) on $env:COMPUTERNAME..."
                    
                    Remove-Item -Path $($usersChild.FullName) -Recurse -Force -ErrorAction Stop
                }
                catch [System.InvalidOperationException] {
                    Write-Verbose "Skipping Removal of $($usersChild.Name) on $env:COMPUTERNAME as Item is Currently In Use..."
                }
            }
        }
    }

    foreach ($profileToDelete in $profilesToDelete) {
        Write-Verbose "Removing Profile $($profileToDelete.LocalPath) & Associated Registry Keys on $env:COMPUTERNAME..."
                
        Remove-WmiObject -InputObject $profileToDelete -ErrorAction Stop
    }

    $finalChildItem = Get-ChildItem -Path "C:\Users" | Select-Object -Property Name, FullName, LastWriteTime
                
    return $finalChildItem
}
## End Remove-UserProfile
## Begin Resolve-ShortURL
Function Resolve-ShortURL {
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
	
    FOREACH ($URL in $ShortUrl) {
        TRY {
            Write-Verbose -Message "$URL - Querying..."
			(Invoke-WebRequest -Uri $URL -MaximumRedirection 0 -ErrorAction Ignore).Headers.Location
        }
        CATCH {
            Write-Error -Message $Error[0].Exception.Message
        }
    }
}
## End Resolve-ShortURL
## Begin Set-ScreenResolution
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
        [Parameter(Mandatory = $true, 
            Position = 0)] 
        [int] 
        $Width, 
 
        [Parameter(Mandatory = $true, 
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
## End Set-ScreenResolution
 
"@ 
 
    Add-Type $pinvokeCode -ErrorAction SilentlyContinue 
    [Resolution.PrmaryScreenResolution]::ChangeResolution($width, $height) 
} 
## End Set-ScreenResolution
## Begin Set-FileTime
Function Set-FileTime {
    param(
        [string[]]$paths,
        [bool]$only_modification = $false,
        [bool]$only_access = $false
    );

    begin {
        Function updateFileSystemInfo([System.IO.FileSystemInfo]$fsInfo) {
            $datetime = get-date
            if ( $only_access ) {
                $fsInfo.LastAccessTime = $datetime
            }
            elseif ( $only_modification ) {
                $fsInfo.LastWriteTime = $datetime
            }
            else {
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
## End Set-FileTime
## Begin Set-PowerShellMemoryTuning
Function Set-PowerShellMemoryTuning {

    param(
        [parameter(
            position = 0,
            mandatory = 1)]
        [ValidateNotNullorEmpty()]
        [ValidateRange(1, 2147483647)]
        [int]
        $memory
    )

    # Test Elevated or not
    $TestElevated = {
        $user = [Security.Principal.WindowsIdentity]::GetCurrent()
        (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

    if (&$TestElevated) {

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
    else {
        Write-Error "This Cmdlet requires Admin right. Please Elevate and try again."
    }

}
## End Set-PowerShellMemoryTuning
## Begin Set-PowerShellWindowTitle
Function Set-PowerShellWindowTitle {
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
## End Set-PowerShellWindowTitle
## Begin Enable-PSScriptBlockLogging
Function Enable-PSScriptBlockLogging {  
    $basePath = "HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging"  
 
    if (-not (Test-Path $basePath)) {  
        $null = New-Item $basePath �Force  
    }
   
    Set-ItemProperty $basePath -Name EnableScriptBlockLogging -Value "1" 
}
## End Enable-PSScriptBlockLogging
## Begin Disable-PSScriptBlockLogging
Function Disable-PSScriptBlockLogging {  
    Remove-Item HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging -Force �Recurse 
}
## End Disable-PSScriptBlockLogging
## Begin Enable-PSScriptBlockInvocationLogging
Function Enable-PSScriptBlockInvocationLogging {
    $basePath = "HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging"  
 
    if (-not (Test-Path $basePath)) {  
        $null = New-Item $basePath �Force  
    }  
 
    Set-ItemProperty $basePath -Name EnableScriptBlockInvocationLogging -Value "1" 
}
## End Enable-PSScriptBlockInvocationLogging
## Begin Enable-PSTranscription
Function Enable-PSTranscription {  
    [CmdletBinding()]  
    param(  
        $OutputDirectory,  
        [Switch] $IncludeInvocationHeader  
    )  
 
    ## Ensure the base path exists  
    $basePath = "HKLM:\Software\Policies\Microsoft\Windows\PowerShell\Transcription" 
    if (-not (Test-Path $basePath)) {  
        $null = New-Item $basePath �Force  
    }

 

    ## Enable transcription  
    Set-ItemProperty $basePath -Name EnableTranscripting -Value 1
 

    ## Set the output directory  
    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey("OutputDirectory")) {  
        Set-ItemProperty $basePath -Name OutputDirectory -Value $OutputDirectory  
    }
 

    ## Set the invocation header  
    if ($IncludeInvocationHeader) {  
        Set-ItemProperty $basePath -Name EnableInvocationHeader -Value 1  
    } 
}
## End Enable-PSTranscription

## Begin Set-ScreenResolution
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
        [Parameter(Mandatory = $true, 
            Position = 0)] 
        [int] 
        $Width, 
 
        [Parameter(Mandatory = $true, 
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
    [Resolution.PrmaryScreenResolution]::ChangeResolution($width, $height) 
} 
## End Set-ScreenResolution
## Begin _SetDocumentProperty
Function _SetDocumentProperty {
    #jeff hicks
    Param(
        [object] $Properties,
        [string] $Name,
        [string] $Value
    )
    #get the property object
    $prop = $properties | ForEach { 
        $propname = $_.GetType().InvokeMember("Name", "GetProperty", $Null, $_, $Null)
        If ($propname -eq $Name) {
            Return $_
        }
    } #ForEach

    #set the value
    $Prop.GetType().InvokeMember("Value", "SetProperty", $Null, $prop, $Value)
}
## End _SetDocumentProperty
## Begin Test-DatePattern
Function Test-DatePattern {
    #http://jdhitsolutions.com/blog/2014/10/powershell-dates-times-and-formats/
    $patterns = "d", "D", "g", "G", "f", "F", "m", "o", "r", "s", "t", "T", "u", "U", "Y", "dd", "MM", "yyyy", "yy", "hh", "mm", "ss", "yyyyMMdd", "yyyyMMddhhmm", "yyyyMMddhhmmss"

    Write-host "It is now $(Get-Date)" -ForegroundColor Green

    foreach ($pattern in $patterns) {

        #create an Object
        [pscustomobject]@{
            Pattern = $pattern
            Syntax  = "Get-Date -format '$pattern'"
            Value   = (Get-Date -Format $pattern)
        }

    } #foreach
    Write-Host "Most patterns are case sensitive" -ForegroundColor Green
}
## End Test-DatePattern
## Begin Test-IsLocalAdministrator
Function Test-IsLocalAdministrator {
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
## End Test-IsLocalAdministrator


## Begin Test-ServerSSLSupport
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
            Host          = $HostName
            Port          = $Port
            SSLv2         = $false
            SSLv3         = $false
            TLSv1_0       = $false
            TLSv1_1       = $false
            TLSv1_2       = $false
            KeyExhange    = $null
            HashAlgorithm = $null
        }
        "ssl2", "ssl3", "tls", "tls11", "tls12" | % {
            $TcpClient = New-Object Net.Sockets.TcpClient
            $TcpClient.Connect($RetValue.Host, $RetValue.Port)
            $SslStream = New-Object Net.Security.SslStream $TcpClient.GetStream()
            $SslStream.ReadTimeout = 15000
            $SslStream.WriteTimeout = 15000
            try {
                $SslStream.AuthenticateAsClient($RetValue.Host, $null, $_, $false)
                $RetValue.KeyExhange = $SslStream.KeyExchangeAlgorithm
                $RetValue.HashAlgorithm = $SslStream.HashAlgorithm
                $status = $true
            }
            catch {
                $status = $false
            }
            switch ($_) {
                "ssl2" { $RetValue.SSLv2 = $status }
                "ssl3" { $RetValue.SSLv3 = $status }
                "tls" { $RetValue.TLSv1_0 = $status }
                "tls11" { $RetValue.TLSv1_1 = $status }
                "tls12" { $RetValue.TLSv1_2 = $status }
            }
            # dispose objects to prevent memory leaks
            $TcpClient.Dispose()
            $SslStream.Dispose()
        }
        $RetValue
    }
}
## End Test-ServerSSLSupport

## Begin Test-SslProtocols
Function Test-SslProtocols {
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
    param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
        $ComputerName,
     
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [int]$Port = 443
    )
    begin {
        $ProtocolNames = [System.Security.Authentication.SslProtocols] | gm -static -MemberType Property | ? { $_.Name -notin @("Default", "None") } | % { $_.Name }
    }
    process {
        $ProtocolStatus = [Ordered]@{}
        $ProtocolStatus.Add("ComputerName", $ComputerName)
        $ProtocolStatus.Add("Port", $Port)
        $ProtocolStatus.Add("KeyLength", $null)
        $ProtocolStatus.Add("SignatureAlgorithm", $null)
     
        $ProtocolNames | % {
            $ProtocolName = $_
            $Socket = New-Object System.Net.Sockets.Socket([System.Net.Sockets.SocketType]::Stream, [System.Net.Sockets.ProtocolType]::Tcp)
            $Socket.Connect($ComputerName, $Port)
            try {
                $NetStream = New-Object System.Net.Sockets.NetworkStream($Socket, $true)
                $SslStream = New-Object System.Net.Security.SslStream($NetStream, $true)
                $SslStream.AuthenticateAsClient($ComputerName, $null, $ProtocolName, $false )
                $RemoteCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]$SslStream.RemoteCertificate
                $ProtocolStatus["KeyLength"] = $RemoteCertificate.PublicKey.Key.KeySize
                $ProtocolStatus["SignatureAlgorithm"] = $RemoteCertificate.PublicKey.Key.SignatureAlgorithm.Split("#")[1]
                $ProtocolStatus.Add($ProtocolName, $true)
            }
            catch {
                $ProtocolStatus.Add($ProtocolName, $false)
            }
            finally {
                $SslStream.Close()
            }
        }
        [PSCustomObject]$ProtocolStatus
    }
}
## End Test-SslProtocols

## Begin Send-Email
Function Send-Email {
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
	
    PROCESS {
        TRY {
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
            IF ($PSBoundParameters['SenderDisplayName']) {
                $SMTPMessage.Sender.DisplayName = $SenderDisplayName
            }
			
            # From Displayname parameter
            IF ($PSBoundParameters['FromDisplayName']) {
                $SMTPMessage.From.DisplayName = $FromDisplayName
            }
			
            # CC Parameter
            IF ($PSBoundParameters['CC']) {
                $SMTPMessage.CC.Add($CC)
            }
			
            # BCC Parameter
            IF ($PSBoundParameters['BCC']) {
                $SMTPMessage.BCC.Add($BCC)
            }
			
            # ReplyToList Parameter
            IF ($PSBoundParameters['ReplyToList']) {
                foreach ($ReplyTo in $ReplyToList) {
                    $SMTPMessage.ReplyToList.Add($ReplyTo)
                }
            }
			
            # Attachement Parameter
            IF ($PSBoundParameters['attachment']) {
                $SMTPattachment = New-Object -TypeName System.Net.Mail.Attachment($attachment)
                $SMTPMessage.Attachments.Add($STMPattachment)
            }
			
            # Delivery Options
            IF ($PSBoundParameters['DeliveryNotificationOptions']) {
                $SMTPMessage.DeliveryNotificationOptions = $DeliveryNotificationOptions
            }
			
            #Create SMTP Client Object
            $SMTPClient = New-Object -TypeName Net.Mail.SmtpClient
            $SMTPClient.Host = $SmtpServer
            $SMTPClient.Port = $Port
			
            # SSL Parameter
            IF ($PSBoundParameters['EnableSSL']) {
                $SMTPClient.EnableSsl = $true
            }
			
            # Credential Paramenter
            #IF (($PSBoundParameters['Username']) -and ($PSBoundParameters['Password']))
            IF ($PSBoundParameters['Credential']) {
                <#
				# Create Credential Object
				$Credentials = New-Object -TypeName System.Net.NetworkCredential
				$Credentials.UserName = $username.Split("@")[0]
				$Credentials.Password = $Password
				#>
				
                # Add the credentials object to the SMTPClient obj
                $SMTPClient.Credentials = $Credential
            }
            IF (-not $PSBoundParameters['Credential']) {
                # Use the current logged user credential
                $SMTPClient.UseDefaultCredentials = $true
            }
			
            # Send the Email
            $SMTPClient.Send($SMTPMessage)
			
        }#TRY
        CATCH {
            Write-Warning -message "[PROCESS] Something wrong happened"
            Write-Warning -Message $Error[0].Exception.Message
        }
    }#Process
    END {
        # Remove Variables
        Remove-Variable -Name SMTPClient -ErrorAction SilentlyContinue
        Remove-Variable -Name Password -ErrorAction SilentlyContinue
    }#END
}
## End Send-Email

## Begin Write-Log
Function Write-Log {
    <#
.SYNOPSIS
    Function to create or append a log file

#>
    [CmdletBinding()]
    Param (
        [Parameter()]
        $Path = "",
        $LogName = "$(Get-Date -f 'yyyyMMdd').log",
        
        [Parameter(Mandatory = $true)]
        $Message = "",

        [Parameter()]
        [ValidateSet('INFORMATIVE', 'WARNING', 'ERROR')]
        $Type = "INFORMATIVE",
        $Category
    )
    BEGIN {
        # Verify if the log already exists, else create it
        IF (-not(Test-Path -Path $(Join-Path -Path $Path -ChildPath $LogName))) {
            New-Item -Path $(Join-Path -Path $Path -ChildPath $LogName) -ItemType file
        }
    
    }
    PROCESS {
        TRY {
            "$(Get-Date -Format yyyyMMdd:HHmmss) [$TYPE] [$category] $Message" | Out-File -FilePath (Join-Path -Path $Path -ChildPath $LogName) -Append
        }
        CATCH {
            Write-Error -Message "Could not write into $(Join-Path -Path $Path -ChildPath $LogName)"
            Write-Error -Message "Last Error:$($error[0].exception.message)"
        }
    }

}
## End Write-Log

## Begin ValidateCompanyName
Function ValidateCompanyName {
    [bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
    If ($xResult) {
        Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
    }
    Else {
        $xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
        If ($xResult) {
            Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
        }
        Else {
            Return ""
        }
    }
}
## End ValidateCompanyName

## Begin ValidateCoverPage
Function ValidateCoverPage {
    Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
    $xArray = ""
	
    Switch ($CultureCode) {
        'ca-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
                    "Integral", "I� (clar)", "I� (fosc)", "L�nia lateral",
                    "Moviment", "Quadr�cula", "Retrospectiu", "Sector (clar)",
                    "Sector (fosc)", "Sem�for", "Visualitzaci� principal", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
                    "Integral", "I� (clar)", "I� (fosc)", "L�nia lateral",
                    "Moviment", "Quadr�cula", "Retrospectiu", "Sector (clar)",
                    "Sector (fosc)", "Sem�for", "Visualitzaci�", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabet", "Anual", "Austin", "Conservador",
                    "Contrast", "Cubicles", "Diplom�tic", "Exposici�",
                    "L�nia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
                    "Perspectiva", "Piles", "Quadr�cula", "Sobri",
                    "Transcendir", "Trencaclosques")
            }
        }

        'da-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Bev�gElse", "Brusen", "Facet", "Filigran", 
                    "Gitter", "Integral", "Ion (lys)", "Ion (m�rk)", 
                    "Retro", "Semafor", "Sidelinje", "Stribet", 
                    "Udsnit (lys)", "Udsnit (m�rk)", "Visningsmaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Bev�gElse", "Brusen", "Ion (lys)", "Filigran",
                    "Retro", "Semafor", "Visningsmaster", "Integral",
                    "Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
                    "Udsnit (m�rk)", "Ion (m�rk)", "Austin")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Bev�gElse", "Moderat", "Perspektiv", "Firkanter",
                    "Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "G�de",
                    "Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
                    "N�lestribet", "�rlig", "Avispapir", "Tradionel")
            }
        }

        'de-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
                    "Geb�ndert", "Integral", "Ion (dunkel)", "Ion (hell)", 
                    "Pfiff", "Randlinie", "Raster", "R�ckblick", 
                    "Segment (dunkel)", "Segment (hell)", "Semaphor", 
                    "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
                    "Raster", "Ion (dunkel)", "Filigran", "R�ckblick", "Pfiff",
                    "ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
                    "Randlinie", "Austin", "Integral", "Facette")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
                    "Herausgestellt", "J�hrlich", "Kacheln", "Kontrast", "Kubistisch",
                    "Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
                    "Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
            }
        }

        'en-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
                    "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
                    "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
                    "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
                    "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
                    "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
            }
        }

        'es-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadr�cula", 
                    "Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
                    "Ion (oscuro)", "L�nea lateral", "Movimiento", "Retrospectiva", 
                    "Sem�foro", "Slice (luz)", "Vista principal", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
                    "Slice (luz)", "Faceta", "Sem�foro", "Retrospectiva", "Cuadr�cula",
                    "Movimiento", "Cortar (oscuro)", "L�nea lateral", "Ion (oscuro)",
                    "Ion (claro)", "Integral", "Con bandas")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
                    "Contraste", "Cuadr�cula", "Cub�culos", "Exposici�n", "L�nea lateral",
                    "Moderno", "Mosaicos", "Movimiento", "Papel peri�dico",
                    "Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
            }
        }

        'fi-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
                    "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
                    "Sektori (vaalea)", "Vaihtuvav�rinen", "ViewMaster", "Austin",
                    "Kuiskaus", "Liike", "Ruudukko", "Sivussa")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
                    "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
                    "Sektori (vaalea)", "Vaihtuvav�rinen", "ViewMaster", "Austin",
                    "Kiehkura", "Liike", "Ruudukko", "Sivussa")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
                    "Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
                    "Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
                    "Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
            }
        }

        'fr-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("� bandes", "Austin", "Facette", "Filigrane", 
                    "Guide", "Int�grale", "Ion (clair)", "Ion (fonc�)", 
                    "Lignes lat�rales", "Quadrillage", "R�trospective", "Secteur (clair)", 
                    "Secteur (fonc�)", "S�maphore", "ViewMaster", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annuel", "Aust�re", "Austin", 
                    "Blocs empil�s", "Classique", "Contraste", "Emplacements de bureau", 
                    "Exposition", "Guide", "Ligne lat�rale", "Moderne", 
                    "Mosa�ques", "Mots crois�s", "Papier journal", "Perspective",
                    "Quadrillage", "Rayures fines", "Transcendant")
            }
        }

        'nb-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
                    "Integral", "Ion (lys)", "Ion (m�rk)", "Retrospekt", "Rutenett",
                    "Sektor (lys)", "Sektor (m�rk)", "Semafor", "Sidelinje", "Stripet",
                    "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabet", "�rlig", "Avistrykk", "Austin", "Avlukker",
                    "BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
                    "Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
                    "Smale striper", "Stabler", "Transcenderende")
            }
        }

        'nl-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
                    "Integraal", "Ion (donker)", "Ion (licht)", "Raster",
                    "Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
                    "Terugblik", "Terzijde", "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
                    "Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
                    "Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
                    "Puzzel", "Raster", "Stapels",
                    "Tegels", "Terzijde")
            }
        }

        'pt-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Anima��o", "Austin", "Em Tiras", "Exibi��o Mestra",
                    "Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
                    "Grade", "Integral", "�on (Claro)", "�on (Escuro)", "Linha Lateral",
                    "Retrospectiva", "Sem�foro")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabeto", "Anima��o", "Anual", "Austero", "Austin", "Baias",
                    "Conservador", "Contraste", "Exposi��o", "Grade", "Ladrilhos",
                    "Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
                    "Quebra-cabe�a", "Transcend")
            }
        }

        'sv-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
                    "Jon (m�rkt)", "Knippe", "Rutn�t", "R�rElse", "Sektor (ljus)", "Sektor (m�rk)",
                    "Semafor", "Sidlinje", "VisaHuvudsida", "�terblick")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabetm�nster", "Austin", "Enkelt", "Exponering", "Konservativt",
                    "Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutn�t",
                    "R�rElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "�rligt",
                    "�verg�ende")
            }
        }

        'zh-'	{
            If ($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ('???', '???', '??', '??', '??',
                    '??(??)', '??(??)', '???', '??', '??(??)',
                    '??(??)', '??', '??', '??', '???',
                    '???')
            }
        }

        Default	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
                    "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
                    "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
                    "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
                    "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
                    "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
            }
        }
    }
	
    If ($xArray -contains $xCP) {
        $xArray = $Null
        Return $True
    }
    Else {
        $xArray = $Null
        Return $False
    }
}
## End ValidateCoverPage

## Begin WMIDateStringToDate
Function WMIDateStringToDate($Bootup) {    
    [System.Management.ManagementDateTimeconverter]::ToDateTime($Bootup)    
} 
## End WMIDateStringToDate


## Begin Force-WSUSCheckin
Function Force-WSUSCheckin($Computer) {
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




## Begin touch
Function touch {
    $args | foreach-object { write-host > $_ } 
}
## End touch
## Begin NPP
Function NPP {
    Start-Process -FilePath "${Env:ProgramFiles(x86)}\Notepad++\Notepad++.exe" 
}#-ArgumentList $args }
## End NPP



## Begin Invoke-VBScript
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
       Written by Simon W�hlin
       http://blog.simonw.se
    #>
    [cmdletbinding(SupportsShouldProcess = $true, ConfirmImpact = 'None', PositionalBinding = $false)]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [ValidateScript({ if (Test-Path $_) { $true }else { Throw "Could not find script: [$_]" } })]
        [String]
        $Path,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Args')]
        [String[]]
        $Argument,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [Switch]
        $Wait
    )
    Begin {
        Write-Verbose -Message 'Locating cscript.exe'
        $cscriptpath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\cscript.exe'
        if (-Not(Test-Path -Path $cscriptpath)) {
            Throw 'cscript.exe not found.'
        }
        Write-Verbose -Message ('cscript.exe found in: {0}' -f $cscriptpath)
    }
    Process {
        Try {
            $ResolvedPath = Resolve-Path -Path $Path
            Write-Verbose -Message ('Processing script: {0}' -f $ResolvedPath)
            if ($PSBoundParameters.ContainsKey('Argument')) {
                $ScriptBlock = [scriptblock]::Create(('& "{0}" "{1}" "{2}"' -f $cscriptpath, $ResolvedPath, ($Argument -join '" "')))
            }
            else {
                $ScriptBlock = $ScriptBlock = [scriptblock]::Create(('& "{0}" "{1}"' -f $cscriptpath, $ResolvedPath))
            }
            Write-Verbose -Message 'Starting script'
            if ($PSCmdlet.ShouldProcess($ResolvedPath, 'Invoke script')) {
                $Job = Start-Job -ScriptBlock $ScriptBlock
                if ($Wait) {
                    $Activity = 'Waiting for script to complete: {0}' -f $ResolvedPath
                    Write-Progress -Activity $Activity -Id 1
                    $i = 1
                    While ($Job.State -eq 'Running') {
                        $WaitTime = (Get-Date) - $Job.PSBeginTime
                        Write-Progress -Activity $Activity -Status "Waited for $($WaitTime.TotalSeconds -as [int]) seconds." -Id 1 -PercentComplete ($i % 100)
                        Start-Sleep -Seconds 1
                        $i++
                    }
                    Write-Progress -Activity $Activity -Status 'Waiting' -Id 1 -Completed
                    $Result = Foreach ($JobInstance in ($Job, $Job.ChildJobs)) {
                        if ($JobInstance.Error -ne $null) {
                            Throw $JobInstance.Error.Exception.Message
                        }
                        else {
                            $JobInstance.Output
                        }
                    }
                    Write-Output -InputObject ($Result -join "`n")
                    Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
                }
                else {
                    Write-Output -InputObject $Job
                }
            }
            Write-Verbose -Message 'Finished processing script'
        }
        Catch {
            Throw
        }
    }
}
## End Invoke-VBScript