## Begin Get-RemoteAppliedGPOs
Function Get-RemoteAppliedGPOs{
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
## End Get-RemoteAppliedGPOs
## Begin Get-GPLink
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
## End Get-GPLink
## Begin Get-GPUnlinked
Function Get-GPUnlinked {
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
## End Get-GPUnlinked
## Begin Get-GPRemote
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
## End Get-GPRemote

$i=0
$j=0

foreach ($Computer in $ComputerName) {

    Write-Progress -Activity "Opening Remote Group Policy..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	#Opens (Remote) Group Policy for specified workstation
	gpedit.msc /gpcomputer: $Computer
    
	}
}
## End Get-GPRemote
## Begin Get-ADGPOReplication
Function Get-ADGPOReplication{
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
## End Get-ADGPOReplication
## Begin GPR
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
## End GPR
## Begin Get-ADGPOReplication
Function Get-ADGPOReplication{
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
## End Get-ADGPOReplication
## Begin Get-ADGPOReplication
Function Get-ADGPOReplication{
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
## End Get-ADGPOReplication
## Begin Set-GPOStatus
Function Set-GPOStatus{
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
## End Set-GPOStatus

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
## End Set-GPOStatus

End {
    Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
} #end
## End Set-GPOStatus

}
## End Set-GPOStatus
## Begin Get-GPRemote
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
## End Get-GPRemote

$i=0
$j=0

foreach ($Computer in $ComputerName) {

    Write-Progress -Activity "Opening Remote Group Policy..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	#Opens (Remote) Group Policy for specified workstation
	gpedit.msc /gpcomputer: $Computer
    
	}
}
## End Get-GPRemote