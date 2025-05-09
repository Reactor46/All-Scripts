#region Internal functions
#helper function for creating generic types
function Get-Type
{
	param(
    	[Parameter(Position=0,Mandatory=$true)]
		[string] $GenericType,
		
		[Parameter(Position=1,Mandatory=$true)]
		[string[]] $T
    )

	$T = $T -as [type[]]
	
	try
	{
		$generic = [type]($GenericType + '`' + $T.Count)
		$generic.MakeGenericType($T)
	}
	catch [Exception]
	{
		throw New-Object System.Exception("Cannot create generic type", $_.Exception)
	}
}

# This function returns the number of trailing zeroes in the input byte
function GetNumberOfTrailingZeroes {
	Param ([byte] $x)
	
	$numOfTrailingZeroes = 0;
	if ( $x -eq 0)
	{
   		return 8
	}
	
	if ( $x % 2 -eq 0)
	{
		$numOfTrailingZeroes ++;
		$numOfTrailingZeroes +=  GetNumberOfTrailingZeroes($x / 2)
	}
	
	return $numOfTrailingZeroes
}

# This function returns the number of non-zero bits in an ip-address
function GetIPAddressPrefixLength {
	Param ([System.Net.IPAddress] $ipAddress)
	
	$byteArray = $ipAddress.GetAddressBytes()
	$numOfTrailingZeroes = 0;
	
	for ($i = $byteArray.Length - 1; $i -ge 0; $i--)
	{
    	$numOfZeroesInByte = GetNumberOfTrailingZeroes($byteArray[$i]);
    	if ($numOfZeroesInByte -eq 0)
		{
			break
		}
    	$numOfTrailingZeroes += $numOfZeroesInByte;
	}
	(($byteArray.Length * 8) - $numOfTrailingZeroes)
}

function IPAddressToDecimal([String]$IPAddress)
{
	[Char[]] $charSplitter = @('.')
	[String[]] $ipData = $IPAddress.Split($charSplitter)
	
	return [Int64](
		([Int32]::Parse($ipData[0]) * [Math]::Pow(2, 24) +
		([Int32]::Parse($ipData[1]) * [Math]::Pow(2, 16) +
		([Int32]::Parse($ipData[2]) * [Math]::Pow(2, 8) +
		([Int32]::Parse($ipData[3]))))))
}

function GetSingleDomainController
{
	param(
		[Parameter(Mandatory=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $DomainControllerName
	)
	
	$rootDSE = [adsi]"LDAP://$script:server/rootDSE"

	$ds = New-Object System.DirectoryServices.DirectorySearcher
	$ds.SearchRoot = "LDAP://$script:server/CN=Sites,$($rootDSE.configurationNamingContext)"
	$ds.Filter = "(&(objectclass=server)(dNSHostName=$DomainControllerName))"

	try
	{
		$serverDn = $ds.FindOne().GetDirectoryEntry().DistinguishedName.Value
		$siteDn = $serverDn.Substring($serverDn.IndexOf("CN=Servers") + 11, $serverDn.IndexOf("CN=Sites") - ($serverDn.IndexOf("CN=Servers") + 12))
		$siteName = $siteDn.Substring(3)
	}
	catch [Exception]
	{
		throw "Could not find domain controller in $($ds.SearchRoot.DistinguishedName): $($_.Exception.Message)"
	}

	try
	{
		$dc = (Get-ADSite -Name $siteName -Server $script:Server).Servers | Where-Object { $_.Name -eq $DomainControllerName }
		if (-not $dc)
		{
			throw "The server $DomainControllerName cannot be found"
		}
		$dc
	}
	catch [Exception]
	{
		throw "Cannot read servers from site $($siteName): ($_.Exception.Message)"
	}
}

function GetADServerObject
{
	[CmdletBinding()]
	param(		
		[Parameter(Mandatory=$true)]
		[string] $Name,
		[Parameter(Mandatory=$true)]
		[string] $Server
	)
		
	if ($Name.Contains("."))
	{
		$Name = $Name.Substring(0, $Name.IndexOf("."))
	}
	
	$rootDSE = [adsi]"LDAP://$server/RootDSE"
	$ds = New-Object System.DirectoryServices.DirectorySearcher
	$ds.SearchRoot = [adsi]"LDAP://$Server/$($rootDSE.configurationNamingContext)"
	$ds.PageSize = 1000
	$ds.Filter = "(&(cn=$Name)(objectClass=server))"
	
	$sr = $ds.FindOne()
	if (-not $sr)
	{
		throw "Server could not be found in container $($ds.SearchRoot)"
	}
	
	$sr.GetDirectoryEntry()
}

function GetADServerNTDSSettingsObject
{
	[CmdletBinding()]
	param(		
		[Parameter(Mandatory=$true)]
		[string] $Name,
		[Parameter(Mandatory=$true)]
		[string] $Server
	)
		
	if ($Name.Contains("."))
	{
		$Name = $Name.Substring(0, $Name.IndexOf("."))
	}
	
	$serverObject = GetADServerObject @PSBoundParameters
	
	$serverNtdsSettingsDN = "CN=NTDS Settings," + $serverObject.distinguishedName
	[adsi]"LDAP://$server/$serverNtdsSettingsDN"	
}

function ConvertTo-AdDateString([DateTime] $date)
{
    $year = $date.Year.ToString()
    $month = $date.Month
    $day = $date.Day
	$hour = $date.Hour
	$minute = $date.Minute
	$second = $date.Second

    $sb = New-Object System.Text.StringBuilder	
    [Void]$sb.Append($year)
	
    if ($month -lt 10){
        [Void]$sb.Append("0")
    }	
    [Void]$sb.Append($month.ToString())
	
    if ($day -lt 10){
        [Void]$sb.Append("0")
    }	
    [Void]$sb.Append($day.ToString())
	
	if ($hour -lt 10){
        [Void]$sb.Append("0");
    }	
    [Void]$sb.Append($hour.ToString())
	
	if ($minute -lt 10){
        [Void]$sb.Append("0")
    }	
    [Void]$sb.Append($minute.ToString())
	
	if ($second -lt 10){
        [Void]$sb.Append("0")
    }	
    [Void]$sb.Append($second.ToString())
	
    [Void]$sb.Append(".0Z")
    return $sb.ToString();
}

function Write-LiveLine
{
	param([Parameter(ValueFromPipeline=$true)][String] $message)

	$str = "`r" + $message
	$padding = $host.UI.RawUI.BufferSize.Width - $str.Length
	$str += ' ' * $padding
	$str += "`r"

	Write-Host $str -NoNewline
}

#region .net types
$typesDefinitions = @'
using System.Collections.Generic;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;

namespace System.DirectoryServices.ActiveDirectory
{
    interface ICast<T>
    {
        T Cast { get; }
    }

    public class ActiveDirectorySchedule2 : ActiveDirectorySchedule
    {
        public string SiteLinkName { get; set; }
        //public Dictionary<DayOfWeek, ScheduleEntry> FormattedSchedule
        public System.Collections.ArrayList FormattedSchedule
        {
            get
            {
                //Dictionary<DayOfWeek, ScheduleEntry> formattedSchedule = new Dictionary<DayOfWeek, ScheduleEntry>();
                System.Collections.ArrayList formattedSchedule = new System.Collections.ArrayList();

                for (int iDay = 0; iDay < this.RawSchedule.GetLength(0); iDay++)
                {
                    ScheduleEntry sw = new ScheduleEntry();
                    for (int iHour = 0; iHour < this.RawSchedule.GetLength(1); iHour++)
                    {
                        bool hxxEnabled = false;
                        for (int iTimeSlice = 0; iTimeSlice < this.RawSchedule.GetLength(2); iTimeSlice++)
                        {
                            if (this.RawSchedule[iDay, iHour, iTimeSlice] == true)
                                hxxEnabled = true;
                        }

                        sw.GetType().GetProperty(string.Format("h{0:00}", iHour)).SetValue(sw, hxxEnabled ? 1 : 0, null);
                        sw.Day = (DayOfWeek)iDay;
                        sw.SiteLinkName = this.SiteLinkName;
                    }
                    //formattedSchedule.Add((System.DayOfWeek)iDay, sw);
                    formattedSchedule.Add(sw);
                }

                return formattedSchedule;
            }
        }

        public ActiveDirectorySchedule2(ActiveDirectorySchedule Schedule, string SiteLinkName)
            : base(Schedule)
        {
            this.SiteLinkName = SiteLinkName;
        }

        public ActiveDirectorySchedule2()
            : base()
        { }

        public void RemoveDailyTimeSlot(HourOfDay fromHour, HourOfDay toHour)
        {
            if (RawSchedule == null)
                RawSchedule = new bool[7, 24, 4];

            bool[, ,] newSchedule = new bool[7, 24, 4];
            newSchedule = RawSchedule;

            for (int d = 0; d < 7; d++)
            {
                for (int fH = (int)fromHour; fH <= (int)toHour; fH++)
                {
                    newSchedule[d, fH, 0] = false;
                    newSchedule[d, fH, 1] = false;
                    newSchedule[d, fH, 2] = false;
                    newSchedule[d, fH, 3] = false;
                }
            }

            RawSchedule = newSchedule;
        }

        public void AddDailyTimeSlot(HourOfDay fromHour, HourOfDay toHour)
        {
            if (RawSchedule == null)
                RawSchedule = new bool[7, 24, 4];

            bool[, ,] newSchedule = new bool[7, 24, 4];
            newSchedule = RawSchedule;

            for (int d = 0; d < 7; d++)
            {
                for (int fH = (int)fromHour; fH <= (int)toHour; fH++)
                {
                    newSchedule[d, fH, 0] = true;
                    newSchedule[d, fH, 1] = true;
                    newSchedule[d, fH, 2] = true;
                    newSchedule[d, fH, 3] = true;
                }
            }

            RawSchedule = newSchedule;
        }

        public void AddTimeSlot(DayOfWeek fromDay, DayOfWeek toDay, HourOfDay fromHour, HourOfDay toHour)
        {
            if (RawSchedule == null)
                RawSchedule = new bool[7, 24, 4];

            bool[, ,] newSchedule = new bool[7, 24, 4];
            newSchedule = RawSchedule;

            for (int d = (int)fromDay; d < (int)toDay; d++)
            {
                for (int fH = (int)fromHour; fH <= (int)toHour; fH++)
                {
                    newSchedule[d, fH, 0] = true;
                    newSchedule[d, fH, 1] = true;
                    newSchedule[d, fH, 2] = true;
                    newSchedule[d, fH, 3] = true;
                }
            }

            RawSchedule = newSchedule;
        }

        public void RemoveTimeSlot(DayOfWeek fromDay, DayOfWeek toDay, HourOfDay fromHour, HourOfDay toHour)
        {
            if (RawSchedule == null)
                RawSchedule = new bool[7, 24, 4];

            bool[, ,] newSchedule = new bool[7, 24, 4];
            newSchedule = RawSchedule;

            for (int d = (int)fromDay; d < (int)toDay; d++)
            {
                for (int fH = (int)fromHour; fH <= (int)toHour; fH++)
                {
                    newSchedule[d, fH, 0] = false;
                    newSchedule[d, fH, 1] = false;
                    newSchedule[d, fH, 2] = false;
                    newSchedule[d, fH, 3] = false;
                }
            }

            RawSchedule = newSchedule;
        }

        public class ScheduleEntry
        {
            public string SiteLinkName { get; set; }
            public DayOfWeek Day { get; set; }
            public int h00 { get; set; }
            public int h01 { get; set; }
            public int h02 { get; set; }
            public int h03 { get; set; }
            public int h04 { get; set; }
            public int h05 { get; set; }
            public int h06 { get; set; }
            public int h07 { get; set; }
            public int h08 { get; set; }
            public int h09 { get; set; }
            public int h10 { get; set; }
            public int h11 { get; set; }
            public int h12 { get; set; }
            public int h13 { get; set; }
            public int h14 { get; set; }
            public int h15 { get; set; }
            public int h16 { get; set; }
            public int h17 { get; set; }
            public int h18 { get; set; }
            public int h19 { get; set; }
            public int h20 { get; set; }
            public int h21 { get; set; }
            public int h22 { get; set; }
            public int h23 { get; set; }
        }
    }

    public class AttributeMetadata2 : ICast<AttributeMetadata>
    {
        protected AttributeMetadata metaData;
        
        public string Name { get { return this.metaData.Name; } }
        public DateTime LastOriginatingChangeTime { get { return this.metaData.LastOriginatingChangeTime; } }
        public Guid LastOriginatingInvocationId { get { return this.metaData.LastOriginatingInvocationId; } }
        public long LocalChangeUsn { get { return this.metaData.LocalChangeUsn; } }
        public long OriginatingChangeUsn { get { return this.metaData.OriginatingChangeUsn; } }
        public string OriginatingServer { get { return this.metaData.OriginatingServer; } }
        public int Version { get { return this.metaData.Version; } }
 
        protected string sourceDomainController;
        public string SourceDomainController
        {
            get { return sourceDomainController; }
            set { sourceDomainController = value; }
        }
 
        public AttributeMetadata2(AttributeMetadata MetaData)
        {
            this.metaData = MetaData;
        }
 
        public static implicit operator AttributeMetadata(AttributeMetadata2 md)
        {
            return md.metaData;
        }
 
        public static implicit operator AttributeMetadata2(AttributeMetadata MetaData)
        {
            return new AttributeMetadata2(new AttributeMetadata2(MetaData));
        }
 
        public override bool Equals(object obj)
        {
            return this.metaData == (AttributeMetadata)obj;
        }
        public override int GetHashCode()
        {
            return this.metaData.GetHashCode();
        }
        public override string ToString()
        {
            return metaData.ToString();
        }
 
        public AttributeMetadata Cast
        {
            get { return this.metaData; }
        }
    }
}

namespace System.DirectoryServices
{
	public class ChangedSearchResult
	{
		public string SourceDomainController { get; set; }
		public string DistinguishedName { get; set; }
		public string Name { get; set; }
		public string UsnChanged { get; set; }
		public string WhenChanged { get; set; }
		public string Path { get; set; }
		
		public string OriginatingServer { get; set; }
		public System.DirectoryServices.ActiveDirectory.AttributeMetadata2[] Metadata { get; set; }
	}
}
'@
#endregion
#endregion

#---------------------------------------------------

#region Remove-ADSite
function Remove-ADSite
{
	[CmdletBinding(
		ConfirmImpact="Low",
		DefaultParameterSetName="Name"
	)]
	
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="Name")]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		
		[Parameter(Mandatory=$false)]
		[string] $Server
	)
 
	begin {
		$script:ctx = $null
		try
		{
			if ($Server)
			{
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
			}
#			else
#			
#			
#			$script:ctx  = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", "msadc430p.ad.vkb.loc")
			
			
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$site = $null
		try
		{
			$site = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::FindByName($script:ctx, $Name)
		}
		catch [Exception]
		{
			Write-Error -Exception $_ -Message "Could not get the site to delete"
			return
		}
		
		try
		{
			$site.Delete()
			Write-Warning "Site $Name was removed"
		}
		catch [Exception]
		{
			Write-Error "Could not remove site $($Name): $($_.Exception.Message)"
		}
    }
	
	end { }
}
#endregion

#region Get-ADSite
function Get-ADSite
{
<#
	.SYNOPSIS
		Gets sites defined in Active Directory
		
	.DESCRIPTION
		The cmdlet returns one or many sites defined in Active Direcotry. You can either get a site 
		
		- by name
		- by IP Address
		- the computer's current site or
		- all sites
		
		
	
	.INPUTS
		System.String
    	The Name of the site
	
	.OUTPUTS
		System.DirectoryServices.ActiveDirectory.ActiveDirectorySite

	.PARAMETER Name
		The name of the site you want to get
		
	.PARAMETER All
		If set all Active Directory Sites defined in the forest are returned.
	
	.PARAMETER IPAddress
		IPAddresses are linked to a site using subnets definitions. When an IPAddress is defined the cmdlet looks for a matching subnet in Active Directory and returns the site that is linked to that subnet.
		
	.PARAMETER Current
		Returns the site of the workstation the cmdlet is running on. It works almost like the IPAddress parameter does by just using the current IP address.
		
	.LINK
		http://gallery.technet.microsoft.com/scriptcenter/780a2272-06f9-4895-827e-9f56bc9272c4
			
	.EXAMPLE
		PS C:\> Get-ADSite -Name Munich

		Name   Location Description Options                       Server Count Subnets
		----   -------- ----------- -------                       ------------ -------
		Munich                      AutoInterSiteTopologyDisabled 16           {192.168.10.0/24}

		Description
		-----------
		This command retreived one specific site defined by the parameter 'Name'.
#>
	[CmdletBinding(
		ConfirmImpact="Low",
		DefaultParameterSetName="Name"
	)]
	
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="Name")]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		
		[Parameter(Mandatory=$true,ParameterSetName="All")]
		[switch] $All,
		
		[Parameter(Mandatory=$true,ParameterSetName="Current")]
		[switch] $Current,
		
		[Parameter(Mandatory=$true,ParameterSetName="ByIPAddress")]
		[System.Net.IPAddress] $IPAddress,
		
		[Parameter(Mandatory=$false)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $Server
	)
 
	begin {
		$script:ctx = $null
		try
		{
			if (-not $Server)
			{
				$script:Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			}
			else
			{
				$script:Server = $Server
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			}
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$site = $null
		try
		{
			switch ($pscmdlet.ParameterSetName)
			{
				"name"
				{
					[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::FindByName($script:ctx, $name)
				}
				"all"
				{
					[System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
				}
				"current"
				{
					[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()
				}
				"byIPAddress"
				{
					$result = New-Object System.Collections.ArrayList
					$subnets = Get-ADSubnet -All
					[Int64] $ipAddressInDecimal = IPAddressToDecimal($IPAddress);
					
					foreach ($subnet in $subnets)
					{
						$splittedSubnetName = $subnet.Name.Split('/');
						
						if ($splittedSubnetName.Length -eq 2)
						{
							[Int32] $subnetMask = [Int32]::Parse($splittedSubnetName[1]);
							[Int32] $numberOfAddresses = [Int32][Math]::Pow(2, (32 - $subnetMask)) - 1
							[string]$ipRange = $splittedSubnetName[0]
							[Int64] $lowIPAddress = IPAddressToDecimal($ipRange)
							[Int64] $highIPAddress = $lowIPAddress + $numberOfAddresses
							[Int64] $totalIPAddressCount = ([Int64][Math]::Pow(2, 31)) - 1

							if (($lowIPAddress -le $ipAddressInDecimal) -and ($ipAddressInDecimal -le $highIPAddress) -and ($numberOfAddresses -le $totalIPAddressCount))
							{
								$result.Add($subnet) | Out-Null
							}
						}
					}
					
					if ($result.Count -gt 1)
					{
						Write-Warning "Multiple Subnets found matching the IP Address"						
					}
					$result
				}
			}
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not get the site / sites"
		}
    }
	
	end { }
}
#endregion

#region New-ADSite
function New-ADSite
{
	[CmdletBinding(ConfirmImpact="Low")]
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $Description,
		[Parameter(Mandatory=$false,Position=1,ValueFromPipelineByPropertyName=$true)]
		[string] $Location,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteOptions] $Options,
		[Parameter(Mandatory=$false)]
		[switch] $PassThru
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$site = $null
		try
		{
			$site = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::FindByName($script:ctx, $Name)
		}
		catch { }
		if ($site)
		{
			Write-Error "Site does already exist"
			return
		}
		
		try
		{
			$site = New-Object System.DirectoryServices.ActiveDirectory.ActiveDirectorySite($script:ctx, $Name)
			if ($Location)
			{
				$site.Location = $Location
			}
			
			if ($Options)
			{
				$Site.Options = $Options
			}
			
			$site.Save()
		
			if (($Description) -and ($Description -ne $site.Description))
			{
				$site.Description = $Description
			}
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not create site: $($_.Exception.Message)"
			return
		}
		
		if ($PassThru)
		{
			$site
		}
    }
	
	end { }
}
#endregion

#region Set-ADSite
function Set-ADSite
{
	[CmdletBinding(ConfirmImpact="Low")]
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $Description,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $Location,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $ISTG,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteOptions] $Options,
		[Parameter(Mandatory=$false)]
		[switch] $PassThru
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$site = $null
		try
		{
			$site = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::FindByName($script:ctx, $Name)
		}
		catch
		{
			Write-Error "Could not read site $($Name): $($_.Exception.Message)"
		}
		
		if (($Location) -and ($Location -ne $site.Location))
		{
			$site.Location = $Location
		}
		
		if (($Options) -and ($Options -ne $site.Options))
		{
			$Site.Options = $Options
		}
		
		if (($Description) -and ($Description -ne $site.Description))
		{
			$site.Description = $Description
		}
		
		try
		{
			$site.Save()
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not set site: $($_.Exception.Message)"
		}
		
		if ($PassThru)
		{
			$site
		}
    }
	
	end { }
}
#endregion

#---------------------------------------------------

#region Test-ADObject
function Test-ADObject
{
	[CmdletBinding(ConfirmImpact="Low")]
    Param (
    	[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
		[string] $DN
	)
	
	process
	{
		if (-not $DN.StartsWith("LDAP://"))
		{
			$DN = "LDAP://$DN"
		}
		
		try
		{
			return [System.DirectoryServices.DirectoryEntry]::Exists($DN)
		}
		catch [Exception]
		{
			return $false
		}	
	}
}
#endregion

#---------------------------------------------------

#region New-ADSubnet
function New-ADSubnet
{
	[CmdletBinding(ConfirmImpact="Low")]
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[ValidateScript({
			if ($_.Contains("/"))
			{
				$subnetIPAddressStr,$prefixLengthStr = $_.Split("/")
				$subnetIPAddress = [System.Net.IPAddress]::Parse($subnetIPAddressStr)
				$specifiedPrefixLength = [int]::Parse($prefixLengthStr)
	          
				$ipAddressPrefixLength = GetIPAddressPrefixLength $subnetIPAddress
				if ($ipAddressPrefixLength -gt $specifiedPrefixLength)
				{
					throw New-Object System.Management.Automation.PSArgumentException("The subnet prefix length you specified is incorrect. Please check the prefix and try again.")
				}
      		}
			else
			{
				$subnetIPAddress = [System.Net.IPAddress]::Parse($_)
				$prefixLength = GetIPAddressPrefixLength $subnetIPAddress
				$Name = $_ + "/" + $prefixLength
			}
			return $true
		})]
		[string] $Name,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $Description,
		[Parameter(Mandatory=$false,Position=1,ValueFromPipelineByPropertyName=$true)]
		[string] $Location,
		[Parameter(Mandatory=$false,Position=2,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if ([string]::IsNullOrEmpty($_))
			{
				return $true
			}
			
			if (Get-ADSite -Name $_)
			{ return $true }
			else
			{
				throw "The site specified could not be found"
			}
		})]
		[string] $Site,
		[Parameter(Mandatory=$false)]
		[switch] $PassThru
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$subnet = $null
		try
		{
			$subnet = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySubnet]::($script:ctx, $Name)
		}
		catch { }
		if ($subnet)
		{
			Write-Error "Subnet does already exist"
			return
		}
		
		try
		{
			$subnet = New-Object System.DirectoryServices.ActiveDirectory.ActiveDirectorySubnet($script:ctx, $Name)
			if ($Location)
			{
				$subnet.Location = $Location
			}
			
			if ($Site)
			{
				$subnet.Site = Get-ADSite -Name $Site
			}
			
			$subnet.Save()
		
			if (($Description) -and ($Description -ne $subnet.Description))
			{
				$subnet.Description = $Description
			}
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not create subnet: $($_.Exception.Message)"
			return
		}
		
		if ($PassThru)
		{
			$subnet
		}
    }
	
	end { }
}
#endregion

#region Get-ADSubnet
function Get-ADSubnet
{
	[CmdletBinding(
		ConfirmImpact="Low",
		DefaultParameterSetName="Name"
	)]
	
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="Name")]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		
		[Parameter(Mandatory=$true,ParameterSetName="All")]
		[switch] $All
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$subnet = $null
		try
		{
			switch ($pscmdlet.ParameterSetName)
			{
				"name"
				{
					[System.DirectoryServices.ActiveDirectory.ActiveDirectorySubnet]::FindByName($script:ctx, $Name)
				}
				"all"
				{
      				$subnetContainerDN = ("CN=Subnets,CN=Sites," + ([adsi]"LDAP://rootDSE").ConfigurationNamingContext)
					$searchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://" + $subnetContainerDN)
					$ds = New-Object System.DirectoryServices.DirectorySearcher
					$ds.SearchRoot = $searchRoot
					$ds.Filter = "(objectCategory=subnet)"
					$ds.PropertiesToLoad.Add("Name") | Out-Null
					
					foreach ($sn in $ds.FindAll())
					{
						[System.DirectoryServices.ActiveDirectory.ActiveDirectorySubnet]::FindByName($script:ctx, $sn.Properties["name"])
					}
				}
			}
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not get the site / sites"
		}
    }
	
	end { }
}
#endregion

#region Remove-ADSubnet
function Remove-ADSubnet
{
	[CmdletBinding(
		ConfirmImpact="Low",
		DefaultParameterSetName="Name"
	)]
	
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="Name")]
		[ValidateNotNullOrEmpty()]
		[string] $Name
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$subnet = $null
		try
		{
			$subnet = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySubnet]::FindByName($script:ctx, $Name)
		}
		catch [Exception]
		{
			Write-Error -Exception $_ -Message "Could not get the subnet to delete"
			return
		}
		
		try
		{
			$subnet.Delete()
			Write-Warning "Subnet $Name was removed"
		}
		catch [Exception]
		{
			Write-Error "Could not remove subnet $($Name): $($_.Exception.Message)"
		}
    }
	
	end { }
}
#endregion

#region Set-ADSubnet
function Set-ADSubnet
{
	[CmdletBinding(ConfirmImpact="Low")]
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[ValidateScript({
			if ($_.Contains("/"))
			{
				$subnetIPAddressStr,$prefixLengthStr = $_.Split("/")
				$subnetIPAddress = [System.Net.IPAddress]::Parse($subnetIPAddressStr)
				$specifiedPrefixLength = [int]::Parse($prefixLengthStr)
	          
				$ipAddressPrefixLength = GetIPAddressPrefixLength $subnetIPAddress
				if ($ipAddressPrefixLength -gt $specifiedPrefixLength)
				{
					throw New-Object System.Management.Automation.PSArgumentException("The subnet prefix length you specified is incorrect. Please check the prefix and try again.")
				}
      		}
			else
			{
				$subnetIPAddress = [System.Net.IPAddress]::Parse($_)
				$prefixLength = GetIPAddressPrefixLength $subnetIPAddress
				$Name = $_ + "/" + $prefixLength
			}
			return $true
		})]
		[string] $Name,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $Description,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $Location,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if ([string]::IsNullOrEmpty($_))
			{
				return $true
			}
			
			if (Get-ADSite -Name $_)
			{ return $true }
			else
			{
				throw "The site specified could not be found"
			}
		})]
		[string] $Site,
		[Parameter(Mandatory=$false)]
		[switch] $PassThru
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$subnet = $null
		try
		{
			$subnet = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySubnet]::FindByName($script:ctx, $Name)
		}
		catch
		{
			throw New-Object System.Exception("Could not read subnet $($Name): $($_.Exception.Message)", $_.Exception)
		}
		
		if ($Location)
		{
			$subnet.Location = $Location
		}
		
		if ($Site)
		{
			$subnet.Site = Get-ADSite -Name $Site
		}
		
		if (($Description) -and ($Description -ne $subnet.Description))
		{
			$subnet.Description = $Description
		}
		
		try
		{
			$subnet.Save()
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not set subnet: $($_.Exception.Message)"
		}
		
		if ($PassThru)
		{
			$subnet
		}
    }
	
	end { }
}
#endregion

#---------------------------------------------------

#region Remove-ADSiteLink
function Remove-ADSiteLink
{
	[CmdletBinding(
		ConfirmImpact="Low",
		DefaultParameterSetName="Name"
	)]
	
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="Name")]
		[ValidateNotNullOrEmpty()]
		[string] $Name
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$siteLink = $null
		try
		{
			$siteLink = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink]::FindByName($script:ctx, $Name)
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not get the site link to delete"
			return
		}
		
		try
		{
			$siteLink.Delete()
			Write-Warning "Site link $Name was removed"
		}
		catch [Exception]
		{
			Write-Error "Could not remove site link $($Name): $($_.Exception.Message)"
		}
    }
	
	end { }
}
#endregion

#region Get-ADSiteLink
function Get-ADSiteLink
{
	[CmdletBinding(
		ConfirmImpact="Low",
		DefaultParameterSetName="Name"
	)]
	
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="Name")]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		
		[Parameter(Mandatory=$true,ParameterSetName="All")]
		[switch] $All		
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$siteLink = $null
		try
		{
			switch ($pscmdlet.ParameterSetName)
			{
				"name"
				{
					try
					{
						try
						{
							$siteLink = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink]::FindByName($script:ctx, $name)
						}
						catch { }
						
						if (-not $siteLink)
						{
							$siteLink = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink]::FindByName($script:ctx, $name, [System.DirectoryServices.ActiveDirectory.ActiveDirectoryTransportType]::Smtp)
						}
					}
					catch [Exception]
					{ 
						Write-Error -Message $_.Message -Exception $_
					}
					
					return $siteLink
				}
				"all"
				{
					$siteLinkContainerDN = ("CN=Inter-Site Transports,CN=Sites," + ([adsi]"LDAP://rootDSE").ConfigurationNamingContext)
					$searchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://" + $siteLinkContainerDN)
					$ds = New-Object System.DirectoryServices.DirectorySearcher
					$ds.SearchRoot = $searchRoot
					$ds.Filter = "(objectCategory=siteLink)"
					$ds.PropertiesToLoad.Add("Name") | Out-Null
					
					foreach ($sl in $ds.FindAll())
					{
						if ($sl.Path.Contains("CN=IP"))
						{
							[System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink]::FindByName($script:ctx, $sl.Properties["name"], [System.DirectoryServices.ActiveDirectory.ActiveDirectoryTransportType]::Rpc)
						}
						elseif ($sl.Path.Contains("CN=SMTP"))
						{
							[System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink]::FindByName($script:ctx, $sl.Properties["name"], [System.DirectoryServices.ActiveDirectory.ActiveDirectoryTransportType]::Smtp)
						}						
					}
				}
			}
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not get the site link / site links"
		}
    }
	
	end { }
}
#endregion

#region New-ADSiteLink
<#
Sample 3
$sites = Get-ADSite -All | Where-Object { $_.Name -ne 'Default-First-Site-Name' }
$mainSite = Get-ADSite -Name Default-First-Site-Name
$sites | ForEach-Object { New-ADSiteLink -Name "$($_.Name)-$($mainSite.Name)" -Description "Automatically created" -Sites $mainSite,$_ -Cost 100 -Interval 15 -NotificationEnabled }
#>
function New-ADSiteLink
{
	[CmdletBinding(ConfirmImpact="Low")]
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $Description,
		[Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidateCount(2, 50)]
		[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite[]] $Sites,
		[Parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$true)]
		[int] $Cost,
		[Parameter(Mandatory=$true,Position=2,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if ($_ -isnot [System.Int32] -and $_ -isnot [System.TimeSpan])
			{ throw "The replication interval must be of type Int32 or TimeSpan" }
			
			if ($_ -is [Int32])
			{
				if ($_ -lt 15)
				{ throw "The replication interval cannot be less than 15 minutes" }
			}
			elseif ($_ -is [System.TimeSpan])
			{
				if ($_.TotalMinutes -lt 15)
				{ throw "The replication interval cannot be less than 15 minutes" }
			}
			return $true
		})]
		$ReplicationInterval,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[switch] $ReciprocalReplicationEnabled,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[switch] $NotificationEnabled,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[switch] $DataCompressionEnabled,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[System.DirectoryServices.ActiveDirectory.ActiveDirectoryTransportType]$TransportType = [System.DirectoryServices.ActiveDirectory.ActiveDirectoryTransportType]::Rpc,
		[Parameter(Mandatory=$false)]
		[switch] $PassThru
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$siteLink = $null
		try
		{
			$siteLink = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink]::FindByName($script:ctx, $Name)
		}
		catch { }
		if ($siteLink)
		{
			Write-Error "Site Link does already exist"
			return
		}
		
		try
		{
			$siteLink = New-Object System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink($script:ctx, $Name, $TransportType)
			$siteLink.Cost = $Cost
			
			if ($ReplicationInterval -is [System.Int32])
			{ $siteLink.ReplicationInterval = New-TimeSpan -Minutes $ReplicationInterval }
			elseif ( $ReplicationInterval -is [System.TimeSpan] )
			{ $siteLink.ReplicationInterval = $ReplicationInterval }
			
			$siteLink.DataCompressionEnabled = $DataCompressionEnabled
			$siteLink.NotificationEnabled = $NotificationEnabled
			$siteLink.ReciprocalReplicationEnabled = $ReciprocalReplicationEnabled
			
			if ($sites)
			{
				foreach ($site in $Sites)
				{
					$siteLink.Sites.Add($site) | Out-Null
				}
			}
			
			$siteLink.Save()
			
			if (($Description) -and ($Description -ne $siteLink.Description))
			{
				$siteLink.Description = $Description
			}
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not create site link: $($_.Exception.Message)"
			return
		}
		
		if ($PassThru)
		{
			$siteLink
		}
    }
	
	end { }
}
#endregion

#region Set-ADSiteLink
function Set-ADSiteLink
{
	[CmdletBinding(ConfirmImpact="Low")]
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string] $Description,
		[Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidateCount(2, 50)]
		[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite[]] $Sites,
		[Parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$true)]
		[int] $Cost,
		[Parameter(Mandatory=$true,Position=2,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if ($_ -isnot [System.Int32] -and $_ -isnot [System.TimeSpan])
			{ throw "The replication interval must be of type Int32 or TimeSpan" }
			
			if ($_ -is [Int32])
			{
				if ($_ -lt 15)
				{ throw "The replication interval cannot be less than 15 minutes" }
			}
			elseif ($_ -is [System.TimeSpan])
			{
				if ($_.TotalMinutes -lt 15)
				{ throw "The replication interval cannot be less than 15 minutes" }
			}
			return $true
		})]
		$ReplicationInterval,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[bool] $ReciprocalReplicationEnabled,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[bool] $NotificationEnabled,
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[bool] $DataCompressionEnabled,
		[Parameter(Mandatory=$false)]
		[switch] $PassThru
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$siteLink = $null
		
		try
		{
			try
			{
				$siteLink = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink]::FindByName($script:ctx, $name)
			}
			catch { }
			
			if (-not $siteLink)
			{
				$siteLink = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink]::FindByName($script:ctx, $name, [System.DirectoryServices.ActiveDirectory.ActiveDirectoryTransportType]::Smtp)
			}
		}
		catch [Exception]
		{ 
			Write-Error -Message $_.Message -Exception $_
		}
		
		if (($Cost) -and ($Cost -ne $siteLink.Cost))
		{
			$siteLink.Cost = $Cost
		}
		
		if ($ReplicationInterval -is [System.Int32])
		{
			if (($ReplicationInterval) -and ($ReplicationInterval -ne $siteLink.ReplicationInterval.TotalMinutes))
			{
				$siteLink.ReplicationInterval = New-TimeSpan -Minutes $ReplicationInterval
			}
		}
		elseif ($ReplicationInterval -is [System.TimeSpan] )
		{
			if ($siteLink.ReplicationInterval -ne $ReplicationInterval)
			{
				$siteLink.ReplicationInterval = $ReplicationInterval
			}
		}
		
		if (($ReciprocalReplicationEnabled -ne $siteLink.ReciprocalReplicationEnabled))
		{
			$siteLink.ReciprocalReplicationEnabled = $ReciprocalReplicationEnabled
		}
		if (($NotificationEnabled -ne $siteLink.NotificationEnabled))
		{
			$siteLink.NotificationEnabled = $NotificationEnabled
		}
		if (($DataCompressionEnabled -ne $siteLink.DataCompressionEnabled))
		{
			$siteLink.DataCompressionEnabled = $DataCompressionEnabled
		}
		
		if (($Description) -and ($Description -ne $site.Description))
		{
			$siteLink.Description = $Description
		}
		
		try
		{
			$siteLink.Save()
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not set site link: $($_.Exception.Message)"
		}
		
		if ($PassThru)
		{
			$siteLink
		}
    }
	
	end { }
}
#endregion

#---------------------------------------------------

#region Get-ADReplicationSchedule
function Get-ADReplicationSchedule
{
	[CmdletBinding(
		ConfirmImpact="Low"
	)]
	
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink] $SiteLink
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		try
		{			
			if ($SiteLink.InterSiteReplicationSchedule)
			{
				$schedule = New-Object System.DirectoryServices.ActiveDirectory.ActiveDirectorySchedule2($SiteLink.InterSiteReplicationSchedule, $siteLink.Name)
				$schedule.FormattedSchedule
			}
			else
			{
				Write-Warning "The site link has no schedule configured"
			}
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not get the site link / site links"
		}
    }
	
	end { }
}
#endregion

#region Reset-ADReplicationSchedule
function Reset-ADReplicationSchedule
{
	[CmdletBinding(
		ConfirmImpact="Low"
	)]
	
	Param (
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[System.DirectoryServices.ActiveDirectory.ActiveDirectorySiteLink] $SiteLink
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		try
		{			
			$schedule = New-Object System.DirectoryServices.ActiveDirectory.ActiveDirectorySchedule2
			$schedule.ResetSchedule()
			$schedule.AddDailyTimeSlot(0, 23)
			$siteLink.InterSiteReplicationSchedule = $schedule
			$siteLink.Save()
			Write-Warning "Inter Site Replication Schedule on site link $($siteLink.Name) was has been resetted (open 24x7)"
		}
		catch [Exception]
		{
			Write-Error -Exception $_.Exception -Message "Could not get the site link / site links"
		}
    }
	
	end { }
}
#endregion

#region Set-ADReplicationSchedule
function Set-ADReplicationSchedule
{
	throw "Not implemented yet"
}
#endregion

#---------------------------------------------------

#region Invoke-KCC
function Invoke-KCC
{
	[CmdletBinding()]
	
	param(
		[Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[Alias("DCName")]
		[string] $DomainControllerName
	)
	
	begin {
		$script:Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
		$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			
		$script:serverObject = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController()
	}
 
	process {
		Write-Host "Connecting to domain controller $DomainControllerName"
		$dc = GetSingleDomainController  -DomainControllerName $DomainControllerName
		
		$dc.CheckReplicationConsistency()
		Write-Host "KCC ran on $DomainControllerName"
    }
	
	end { }
}
#endregion Invoke-KCC

#region Get-ADReplicationConnection
function Get-ADReplicationConnection
{
	[CmdletBinding()]
	
	param(
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInSite",ValueFromPipelineByPropertyName=$true)]
		[string] $SiteName,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInForest")]
		[switch] $AllDCsInForest,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInDomain")]
		[switch] $AllDCsInDomain,
		
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $DomainName,
		
		[Parameter(Position=0,Mandatory=$false,ValueFromPipelineByPropertyName=$true,ParameterSetName="DcByName")]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[Alias("DCName")]
		[string[]] $DomainControllerName,
		
		[Parameter(Mandatory=$false)]
		[switch] $ErrorsOnly,
		
		[Parameter(Mandatory=$false)]
		[switch] $ShowOutboundConnections,
		
		[Parameter(Mandatory=$false)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $Server
	)
	
	begin {
		$global:tempDestinationServer = $null
		$script:ctx = $null
		try
		{
			if (-not $Server)
			{
				$script:Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			}
			else
			{
				$script:Server = $Server
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			}
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$dcList = New-Object System.Collections.ArrayList
		
		if ($pscmdlet.ParameterSetName -eq "AllDCsInForest")
		{
			Write-Warning "Creating a list of domain controllers for the whole forest. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
			foreach ($domain in $forest.Domains)
			{
				foreach ($DomainController in $domain.DomainControllers)
				{
					[Void]$dcList.Add($DomainController)
				}
			}
		}
		
		if ($pscmdlet.ParameterSetName -eq "AllDCsInDomain")
		{
			if ($DomainName)
			{
				$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainName)
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
			}
			else
			{
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
			}
			Write-Warning "Creating a list of domain controllers for the whole domain '$domain'. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			
			foreach ($DomainController in $domain.DomainControllers)
			{
				[Void]$dcList.Add($DomainController)
			}
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInSite")
		{
			if ($DomainName)
			{
				$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainName)
				$dcList.AddRange([System.DirectoryServices.ActiveDirectory.DomainController]::FindAll($ctx, $SiteName))
			}
			else
			{
				$site = Get-ADSite -Name $SiteName
				$dcList.AddRange($site.Servers)
			}
		}
		elseif ($pscmdlet.ParameterSetName -eq "DcByName")
		{
			foreach ($dc in $DomainControllerName)
			{
				[Void]$dcList.Add((GetSingleDomainController -DomainControllerName $dc))
			}
		}
		
		$connections = New-Object (Get-Type -GenericType System.Collections.Generic.List -T System.Management.Automation.PSObject)
		foreach ($dc in $dcList)
		{
			if (-not (Test-Connection -ComputerName $dc.Name -Count 1 -ErrorAction SilentlyContinue))
			{
				Write-Warning "Domain Controller $($dc.Name) could not be reached"
				continue
			}
			
			foreach ($connection in $dc.InboundConnections)
			{
				$connection = $connection | Add-Member -MemberType NoteProperty -Name ReadFromServer -Value $dc.Name -PassThru
				$connections.Add([psobject]$connection)
			}
			
			if ($ShowOutboundConnections)
			{
				foreach ($connection in $dc.OutboundConnections)
				{
					$connection = $connection | Add-Member -MemberType NoteProperty -Name ReadFromServer -Value $dc.Name -PassThru
					$connections.Add([psobject]$connection)
				}
			}
		}
		
		$connections
    }
	
	end { }
}
#endregion Get-ADReplicationConnection

#region New-ADReplicationConnection
function New-ADReplicationConnection
{
	[CmdletBinding()]
	
	param(
		[Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[Alias("DCName")]
		[string] $DomainControllerName,
		
		[Parameter(ValueFromPipelineByPropertyName=$true)]
		[Alias("Name")]
		[string] $ConnectionName,
		
		[Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $SourceServer,
		
		[Parameter(Mandatory=$false)]
		[switch] $StartKCC,
		
		[Parameter(ValueFromPipelineByPropertyName=$true)]
				[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $Server
	)
	
	begin {
		$script:ctx = $null
		try
		{
			if (-not $Server)
			{
				$script:Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
				
				$script:serverObject = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController()
			}
			else
			{
				$script:Server = $Server
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
				
				$script:serverObject = GetSingleDomainController -DomainControllerName $Server
			}
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		try
		{
			if ($ConnectionName -eq [string]::Empty) { $ConnectionName = $SourceServer }
			
			Write-Host "Connecting to domain controller $DomainControllerName"
			$ntds = GetADServerNTDSSettingsObject -Name $DomainControllerName -Server $Script:Server
			$fromServer = GetADServerNTDSSettingsObject -Name $SourceServer -Server $script:Server
			
			$connection = $ntds.psbase.Children.Add("CN=$ConnectionName", "nTDSConnection")
			$connection.psbase.Properties["fromServer"].Value = $fromServer.distinguishedName.Value
			$connection.psbase.Properties["Options"].Value = 0
			$connection.psbase.Properties["enabledConnection"].Value = $true
			$connection.psbase.CommitChanges()
			Write-Host "Connection $($connection.Name) created on $DomainControllerName"
		}
		catch [Exception]
		{
			Write-Error -Message "Error creating connection on $($DomainControllerName): $($_.Exception.Message)" -Exception $_.Exception
		}
    }
	
	end
	{
		if ($StartKCC)
		{
			Write-Host "Starting KCC on domain controller $($script:serverObject.Name)"
			$script:serverObject.CheckReplicationConsistency()
		}
	}
}
#endregion New-ADReplicationConnection

#region Remove-ADReplicationConnection
function Remove-ADReplicationConnection
{
	[CmdletBinding()]
	
	param(
		[Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[Alias("DCName")]
		[string] $DomainControllerName,
		
		[Parameter(Mandatory=$true,ParameterSetName="ByName")]
		[string] $ConnectionName,
		
		[Parameter(Mandatory=$false)]
		[switch] $StartKCC,
		
		[Parameter(Mandatory=$true,ParameterSetName="All")]
		[switch] $All,
		
		[Parameter(ValueFromPipelineByPropertyName=$true)]
				[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $Server
	)
	
	begin {
		$script:ctx = $null
		try
		{
			if (-not $Server)
			{
				$script:Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
				
				$script:serverObject = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController()
			}
			else
			{
				$script:Server = $Server
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
				
				$script:serverObject = GetSingleDomainController -DomainControllerName $Server
			}
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		Write-Host "Connecting to domain controller $DomainControllerName"
		$ntds = GetADServerNTDSSettingsObject -Name $DomainControllerName -Server $script:Server

		try
		{
			foreach ($connection in $ntds.Children)
			{
				if ($ConnectionName)
				{
					if ($connection.Name -eq $ConnectionName)
					{
						Write-Warning "Deleting connection $($connection.Name)$DomainControllerName"
						$connection.DeleteTree()
					}
				}
				else
				{
					Write-Warning "Deleting connection $($connection.Name) on $DomainControllerName"
					$connection.DeleteTree()
				}
			}
		}
		catch [Exception]
		{
			Write-Error -Message "Error deleting connection $($connection.Name) on $DomainControllerName" -Exception $_.Exception
		}
		
		if ($StartKCC)
		{
			Write-Host "Starting KCC on domain controller $($script:serverObject.Name)"
			$script:serverObject.CheckReplicationConsistency()
		}
    }
	
	end { }
}
#endregion Remove-ADReplicationConnection

#region Get-ADReplicationLink
function Get-ADReplicationLink
{
	[CmdletBinding()]
	
	param(		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInSite",ValueFromPipelineByPropertyName=$true)]
		[string] $SiteName,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInForest")]
		[switch] $AllDCsInForest,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInDomain")]
		[switch] $AllDCsInDomain,
		
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $DomainName,
		
		[Parameter(Position=0,Mandatory=$false,ValueFromPipelineByPropertyName=$true,ParameterSetName="DcByName")]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[Alias("DCName")]
		[string[]] $DomainControllerName,
		
		[Parameter(Position=0,Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[Alias("NC")]
		[string] $NamingContext,
		
		[Parameter(Mandatory=$false)]
		[switch] $ErrorsOnly,
		
		[Parameter(Mandatory=$false)]
		[string] $Server
	)
	
	begin {
		$global:tempDestinationServer = $null
		$script:ctx = $null
		try
		{
			if (-not $Server)
			{
				$script:Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
				
				$script:serverObject = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController()
			}
			else
			{
				$script:Server = $Server
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
				
				$script:serverObject = GetSingleDomainController -DomainControllerName $Server
			}
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$dcList = New-Object System.Collections.ArrayList
		
		if ($pscmdlet.ParameterSetName -eq "AllDCsInForest")
		{
			Write-Warning "Creating a list of domain controllers for the whole forest. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
			Write-Host "Discovering domains and domain controllers in the forest $forest" -ForegroundColor DarkGreen
			Write-Host "	New Domain = '#'          New Domain Controller = '.'" -ForegroundColor DarkGreen
			Write-Host
			foreach ($domain in $forest.Domains)
			{
				Write-Host "#" -NoNewline
				foreach ($DomainController in $domain.DomainControllers)
				{
					Write-Host '.' -NoNewline
					[Void]$dcList.Add($DomainController)
				}
			}
			Write-Host
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInDomain")
		{
			if ($DomainName)
			{
				$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainName)
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
			}
			else
			{
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
			}
			Write-Warning "Creating a list of domain controllers for the whole domain '$domain'. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			Write-Host "Discovering domains and domain controllers in the forest $forest" -ForegroundColor DarkGreen
			Write-Host
			foreach ($DomainController in $domain.DomainControllers)
			{
				Write-Host '.' -NoNewline
				[Void]$dcList.Add($DomainController)
			}
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInSite")
		{
			if ($DomainName)
			{
				$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainName)
				$dcList.AddRange([System.DirectoryServices.ActiveDirectory.DomainController]::FindAll($ctx, $SiteName))
			}
			else
			{
				$site = Get-ADSite -Name $SiteName
				$dcList.AddRange($site.Servers)
			}
		}
		elseif ($pscmdlet.ParameterSetName -eq "DcByName")
		{
			foreach ($DomainController in $DomainControllerName)
			{
				[Void]$dcList.Add((GetSingleDomainController -DomainControllerName $DomainController))
			}
		}
		
		$repInfo = New-Object (Get-Type -GenericType System.Collections.Generic.List -T System.Management.Automation.PSObject)
		foreach ($dc in $dcList)
		{
			Write-Host "Connecting to domain controller $dc" -ForegroundColor DarkGreen
			if (-not (Test-Connection -ComputerName $dc.Name -Count 1 -ErrorAction SilentlyContinue))
			{
				Write-Warning "Domain Controller $($dc.Name) could not be reached"
				continue
			}
			
			foreach ($partition in $dc.Partitions)
			{
				if ($NamingContext)
				{
					if ($partition -ne $NamingContext)
					{
						continue
					}
				}
				
				$repNeighbors = $dc.GetReplicationNeighbors($partition)
				if ($ErrorsOnly)
				{
					$repNeighbors = $repNeighbors | Where-Object { $_.LastSyncResult -ne 0 }
				}
				if ($repNeighbors.Count -eq 0 -or $repNeighbors -eq $null)
				{
					continue
				}
				
				foreach ($repNeighbor in $repNeighbors)
				{
					$repNeighbor = $repNeighbor | Add-Member -MemberType NoteProperty -Name DestinationServer -Value $dc.Name -PassThru
					$repNeighbor = $repNeighbor | Add-Member -MemberType NoteProperty -Name DestinationServerAndNamingContext -Value "$($dc.Name) $($repNeighbor.PartitionName)" -PassThru					
					$repInfo.Add([psobject]$repNeighbor)
				}
			}
		}
		
		$repInfo
    }
	
	end { }
}
#endregion

#region Get-ReplicationMetadata
function Get-ADReplicationMetadata
{
	[cmdletBinding()]
	param(
		[Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[Alias("DN")]
#		[ValidateScript({
#			if (-not (Test-ADObject -DN $_))
#			{
#				throw "The object $_ could not be found"
#			}
#			return $true
#		})]
		[string] $DistinguishedName,
	
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInSite",ValueFromPipelineByPropertyName=$true)]
		[string] $SiteName,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInForest")]
		[switch] $AllDCsInForest,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInDomain")]
		[switch] $AllDCsInDomain,
		
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $DomainName,
		
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,ParameterSetName="DcByName")]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[Alias("DCName")]
		[string[]] $DomainControllerName,
		
		[Parameter(Mandatory=$false)]
		[string[]] $Attributes,
		
		[Parameter(Mandatory=$false)]
		[string] $Server
	)
 
	begin {
		$script:ctx = $null
		try
		{
			$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest")
			if (-not $Server)
			{
				$Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
			}
			$script:server = $Server
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$dcList = New-Object System.Collections.ArrayList
		
		if ($pscmdlet.ParameterSetName -eq "AllDCsInForest")
		{
			Write-Warning "Creating a list of domain controllers for the whole forest. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
			foreach ($domain in $forest.Domains)
			{
				foreach ($dc in $domain.DomainControllers)
				{
					[Void]$dcList.Add($dc)
				}
			}
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInDomain")
		{
			Write-Warning "Creating a list of domain controllers for the whole domain '$DomainName'. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			if ($DomainName)
			{
				$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainName)
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
			}
			else
			{
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
			}
			
			foreach ($DomainController in $domain.DomainControllers)
			{
				[void]$dcList.Add($DomainController)
			}
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInSite")
		{
			if ($DomainName)
			{
				$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainName)
				$dcList.AddRange([System.DirectoryServices.ActiveDirectory.DomainController]::FindAll($ctx, $SiteName))
			}
			else
			{
				$site = Get-ADSite -Name $SiteName
				$dcList.AddRange($site.Servers)
			}
		}
		elseif ($pscmdlet.ParameterSetName -eq "DcByName")
		{
			foreach ($dc in $DomainControllerName)
			{
				[Void]$dcList.Add((GetSingleDomainController -DomainControllerName $dc))
			}
		}
		
		$attributeMetaDataList = New-Object System.Collections.ArrayList
		
		foreach ($dc in $dcList)
		{
			if (-not (Test-Connection -ComputerName $dc.Name -Count 1 -ErrorAction SilentlyContinue))
			{
				Write-Warning "Domain Controller $($dc.Name) could not be reached"
				continue
			}
			
			$metaData = $dc.GetReplicationMetadata($DistinguishedName)
				
			foreach ($attributeName in $metaData.AttributeNames)
			{
				if ($Attributes)
				{
					if ($Attributes -notcontains $attributeName)
					{
						continue
					}
				}
				
				$attributeMetaData = [System.DirectoryServices.ActiveDirectory.AttributeMetadata2]$metaData.Item("$attributeName")
				$attributeMetaData.SourceDomainController = $dc.Name
				[Void]$attributeMetaDataList.Add($attributeMetaData)
			}
		}
		
		$attributeMetaDataList | Sort-Object -Property Name
    }
	
	end { }
}
#endregion Get-ReplicationMetadata

#region Get-ADReplicationQueue
function Get-ADReplicationQueue
{
<#
	.SYNOPSIS
		Get the replication queue of a domain controller
		
	.DESCRIPTION
		Get the replication queue of a domain controller. If prints out all outstanding replication items. If only interested in a short summary use the SuppressDetails switch.
		
		The command can target one domain controller as well as all domain controllers in a site, domain or the entire forest.
		
		Retreiving the queue of a whole domain or forest can take a long time.
		
	.INPUTS
		System.String or Object
		Name of the domain controller, site, domain or forest. This data can also be extracted from an object if the property names match the parameter names.
	
	.OUTPUTS
		System.DirectoryServices.ActiveDirectory.ReplicationOperation
		
		If SuppressDetails is used:	Selected.System.Management.Automation.PSCustomObject

	.PARAMETER SiteName
		Retreives the queue from all domain controllers in the site.
		
	.PARAMETER AllDCsInDomain
		Retreives the queue from all domain controllers in the current domain or the domain specified in the DomainName parameter.
		
	.PARAMETER AllDCsInForest
		Retreives the queue from all domain controllers the current forest.
	
	.PARAMETER DomainName
		If AllDCsInDomain is used this parameter specifies the domain to get the date from.
		
	.PARAMETER DomainControllerName
		Reads the replication queue of this domain controller.
	
	.PARAMETER SuppressDetails
		This command returns the replication operations found in the replication queue. If you are just interested in the number of items in the queu, use that switch.
		
	.LINK
		http://gallery.technet.microsoft.com/scriptcenter/780a2272-06f9-4895-827e-9f56bc9272c4
			
	.EXAMPLE
		PS C:\> Get-ADReplicationQueue -DomainControllerName F3DC1.f3.net -SuppressDetails

		DomainController                                                              QueueLength
		----------------                                                              -----------
		F3DC1.f3.net		                                                                  230

		Description
		-----------
		This command retreived the replication queue from one domain controller.
		
	.EXAMPLE
		PS C:\> Get-ADReplicationQueue -SiteName Hub -SuppressDetails

		DomainController                                                              QueueLength
		----------------                                                              -----------
		The replication queue on DC401P.F3.NET is empty
		DC402P.F3.NET                             		     	                              154
		The replication queue on DC405P.F3.NET is empty
		The replication queue on DC403P.F3.NET is empty
		DC404P.F3.NET                                   	                                   74
		The replication queue on DC420P.F3.NET is empty
		The replication queue on DC421P.F3.NET is empty
		The replication queue on DC422P.F3.NET is empty
		DC423P.F3.NET                                  		                                 1130
		The replication queue on DC424P.F3.NET is empty
		DC425P.F3.NET                                         		                           38
		The replication queue on DC426P.F3.NET is empty
		DC427P.F3.NET                                                              			  168
		The replication queue on DC428P.F3.NET is empty
		The replication queue on DC429P.F3.NET is empty
		The replication queue on DC430P.F3.NET is empty

		Description
		-----------
		This command retreived the replication queue from one domain controllers om the site 'Hub'
#>
	[cmdletBinding()]
	param(		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInSite",ValueFromPipelineByPropertyName=$true)]
		[string] $SiteName,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInForest")]
		[switch] $AllDCsInForest,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInDomain")]
		[switch] $AllDCsInDomain,
		
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $DomainName,
		
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,ParameterSetName="DcByName")]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[Alias("DCName")]
		[string[]] $DomainControllerName,
		
		[Parameter(Mandatory=$false)]
		[switch] $SuppressDetails,
		
		[Parameter(Mandatory=$false)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $Server
	)
 
	begin {
		$script:ctx = $null
		try
		{
			if (-not $Server)
			{
				$script:Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			}
			else
			{
				$script:Server = $Server
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			}
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
 
	process {
		$dcList = New-Object System.Collections.ArrayList
		
		if ($pscmdlet.ParameterSetName -eq "AllDCsInForest")
		{
			Write-Warning "Creating a list of domain controllers for the whole forest. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
			foreach ($domain in $forest.Domains)
			{
				foreach ($dc in $domain.DomainControllers)
				{
					[Void]$dcList.Add($dc)
				}
			}
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInDomain")
		{
			Write-Warning "Creating a list of domain controllers for the whole domain '$DomainName'. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			if ($DomainName)
			{
				$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainName)
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
			}
			else
			{
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
			}
			
			foreach ($DomainController in $domain.DomainControllers)
			{
				$dcList.Add($dc)
			}
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInSite")
		{
			if ($DomainName)
			{
				$dcList.AddRange([System.DirectoryServices.ActiveDirectory.DomainController]::FindAll($script:ctx, $SiteName))
			}
			else
			{
				$site = Get-ADSite -Name $SiteName
				$dcList.AddRange($site.Servers)
			}
		}
		elseif ($pscmdlet.ParameterSetName -eq "DcByName")
		{
			foreach ($dc in $DomainControllerName)
			{
				[Void]$dcList.Add((GetSingleDomainController -DomainControllerName $dc))
			}
		}
		
		foreach ($dc in $dcList)
		{
			if (-not (Test-Connection -ComputerName $dc.Name -Count 1 -ErrorAction SilentlyContinue))
			{
				Write-Warning "Domain Controller $($dc.Name) could not be reached"
				continue
			}
			
			$repInfo = $dc.GetReplicationOperationInformation().PendingOperations
			if ($repInfo.Count -gt 0)
			{				
				if ($SuppressDetails)
				{
					$summaryObject = New-Object PSObject | Select-Object -Property DomainController,QueueLength
					$summaryObject.DomainController = $dc.Name
					$summaryObject.QueueLength = $repInfo.Count
					$summaryObject
				}
				else
				{
					$repInfo | Add-Member -MemberType NoteProperty -Name DomainController -Value $dc.Name -PassThru
				}
			}
			else
			{
				Write-Host "The replication queue on $($dc.Name) is empty"				
			}
		}
    }
	
	end { }
}
#endregion Get-ADReplicationQueue

#region Get-ADLastChanges
function Get-ADLastChanges
{
	[cmdletBinding()]
	param(		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInSite",ValueFromPipelineByPropertyName=$true)]
		[string] $SiteName,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInForest")]
		[switch] $AllDCsInForest,
		
		[Parameter(Mandatory=$true,ParameterSetName="AllDCsInDomain")]
		[switch] $AllDCsInDomain,
		
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $DomainName,
		
		[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,ParameterSetName="DcByName")]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[Alias("DCName")]
		[string[]] $DomainControllerName,
		
		[Parameter(Mandatory=$false)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string] $Server,
		
		[Parameter(Mandatory=$true)]		
		[object] $Filter,
		
		[Parameter(Mandatory=$false)]
		[ValidateScript({
			Write-Warning "Reading the metadata of each object changed can take a minute"
			return $true
		})]
		[switch] $IncludeMetadata
	)
	
	begin {
		$script:ctx = $null
		try
		{
			if (-not $Server)
			{
				$script:Server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			}
			else
			{
				$script:Server = $Server
				$script:ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer", $script:Server)
			}
		}
		catch [Exception]
		{
			Write-Error $_
			return
		}
	}
	
	process {
		$dcList = New-Object System.Collections.ArrayList
		
		if ($pscmdlet.ParameterSetName -eq "AllDCsInForest")
		{
			Write-Warning "Creating a list of domain controllers for the whole forest. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
			foreach ($domain in $forest.Domains)
			{
				foreach ($dc in $domain.DomainControllers)
				{
					[Void]$dcList.Add($dc)
				}
			}
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInDomain")
		{
			Write-Warning "Creating a list of domain controllers for the whole domain '$DomainName'. Depending on the connectivity and the number of domain controllers this can take some minutes or hours"
			if ($DomainName)
			{
				$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainName)
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
			}
			else
			{
				$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
			}
			
			foreach ($DomainController in $domain.DomainControllers)
			{
				$dcList.Add($dc)
			}
		}
		if ($pscmdlet.ParameterSetName -eq "AllDCsInSite")
		{
			if ($DomainName)
			{
				$dcList.AddRange([System.DirectoryServices.ActiveDirectory.DomainController]::FindAll($script:ctx, $SiteName))
			}
			else
			{
				$site = Get-ADSite -Name $SiteName
				$dcList.AddRange($site.Servers)
			}
		}
		elseif ($pscmdlet.ParameterSetName -eq "DcByName")
		{
			foreach ($dc in $DomainControllerName)
			{
				[Void]$dcList.Add((GetSingleDomainController -DomainControllerName $dc))
			}
		}
		
		foreach ($dc in $dcList)
		{
			$forestDN = (Get-ADForest).RootDomain.GetDirectoryEntry().DistinguishedName

			if ($dc.IsGlobalCatalog())
			{
				$searchRoot = "GC://$($dc.Name)/$forestDN"
				
				$ds = New-Object System.DirectoryServices.DirectorySearcher
				$ds.SearchRoot = $searchRoot
				
				if ($Filter -is [int])
				{
					Write-Host ("highestCommittedUSN on DC {0} is {1}, searching for USNs higher than {2}" -f $dc.Name, $dc.HighestCommittedUsn, ($dc.HighestCommittedUsn - $Filter))
					$ds.Filter = "(&(objectclass=*)(usnChanged>=$($dc.HighestCommittedUsn - $Filter)))"
				}
				elseif ($Filter -is [datetime])
				{
					$adDateTime = ConvertTo-AdDateString $Filter
					$ds.Filter = "(&(objectclass=*)(whenChanged>=$adDateTime))"
					Write-Host ("Searching objects on {0} with whenChanged greater {1} ({2})" -f $dc.Name, $Filter, $adDateTime)
				}
				else
				{
					throw "Invalid object used for parameter Filter. Use either an integer for filtering by USN or a DateTime if you want to filter for whenChanged"
				}
				
				$ds.PropertiesToLoad.AddRange(("distinguishedName","usnChanged","Name","LocalChangeUsn","whenChanged"))
				$ds.Tombstone = $true
				$ds.PageSize = 1000
				$src = $ds.FindAll()

				$srcEnumerator = $src.GetEnumerator()
				while ($srcEnumerator.MoveNext())
				{
					$sr = [System.DirectoryServices.SearchResult]$srcEnumerator.Current
					
					$changedSr = New-Object System.DirectoryServices.ChangedSearchResult
					$changedSr.SourceDomainController = $dc.Name
					$changedSr.DistinguishedName = $sr.Properties.Item("distinguishedName")
					$changedSr.Path = $sr.Path
					$changedSr.Name = $sr.Properties.Item("Name")
					$changedSr.UsnChanged = $sr.Properties.Item("usnChanged")
					$changedSr.WhenChanged = $sr.Properties.Item("whenChanged")
					
					if ($IncludeMetadata)
					{
						Write-Host "." -NoNewline
						$changedSr.Metadata = Get-ADReplicationMetadata -DomainControllerName $dc.Name -DistinguishedName $changedSr.DistinguishedName |
							Sort-Object -Property LocalChangeUsn -Descending
						$changedSr.OriginatingServer = @($changedSr.Metadata | 
							Where-Object { ($_.Name -ne "cn" -and $_.Name -ne "ou" -and $_.Name -ne "dc") -and $_.LocalChangeUsn -eq $changedSr.UsnChanged })[0].OriginatingServer
					}
						
					$changedSr
				}
			}
			else
			{
				foreach ($partition in $dc.Partitions)
				{
					$searchRoot = "LDAP://$($dc.Name)/$partition"
					
					$ds = New-Object System.DirectoryServices.DirectorySearcher
					$ds.SearchRoot = $searchRoot
					
					if ($Filter -is [int])
					{
						Write-Host ("highestCommittedUSN on DC {0} is {1}, searching for USNs higher than {2}" -f $dc.Name, $dc.HighestCommittedUsn, ($dc.HighestCommittedUsn - $Filter))
						$ds.Filter = "(&(objectclass=*)(usnChanged>=$($dc.HighestCommittedUsn - $Filter)))"
					}
					elseif ($Filter -is [datetime])
					{
						$adDateTime = ConvertTo-AdDateString $Filter
						$ds.Filter = "(&(objectclass=*)(whenChanged>=$adDateTime))"
						Write-Host ("Searching objects on {0} with whenChanged greater {1} ({2})" -f $dc.Name, $Filter, $adDateTime)
					}
					else
					{
						throw "Invalid object used for parameter Filter. Use either an integer for filtering by USN or a DateTime if you want to filter for whenChanged"
					}
					
					$ds.PropertiesToLoad.AddRange(("distinguishedName","usnChanged","Name","LocalChangeUsn","whenChanged"))
					$src = $ds.FindAll()

					$srcEnumerator = $src.GetEnumerator()
					while ($srcEnumerator.MoveNext())
					{
						$sr = [System.DirectoryServices.SearchResult]$srcEnumerator.Current
						
						$changedSr = New-Object System.DirectoryServices.ChangedSearchResult
						$changedSr.SourceDomainController = $dc.Name
						$changedSr.DistinguishedName = $sr.Properties.Item("distinguishedName")
						$changedSr.Path = $sr.Path
						$changedSr.Name = $sr.Properties.Item("Name")
						$changedSr.UsnChanged = $sr.Properties.Item("usnChanged")
						$changedSr.WhenChanged = $sr.Properties.Item("whenChanged")
						
						if ($IncludeMetadata)
						{
							Write-Host "." -NoNewline
							$changedSr.Metadata = Get-ADReplicationMetadata -DomainControllerName $dc.Name -DistinguishedName $changedSr.DistinguishedName |
								Sort-Object -Property LocalChangeUsn -Descending
							$changedSr.OriginatingServer = @($changedSr.Metadata | 
								Where-Object { ($_.Name -ne "cn" -and $_.Name -ne "ou" -and $_.Name -ne "dc") -and $_.LocalChangeUsn -eq $changedSr.UsnChanged })[0].OriginatingServer
						}
							
						$changedSr
					}
				}
			}
			Write-Host
		}
	}
	
	end { }
}
#endregion Get-ADLastChanges

#---------------------------------------------------

#region Get-ADForest
function Get-ADForest
{
	[cmdletBinding()]
	param(		
		[Parameter(Mandatory=$false)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string]$Name
	)
	
	process
	{
		if ($Name)
		{
			$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $Name)		
			[System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ctx)
		}
		else
		{
			[System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
		}
	}
}
#endregion Get-ADForest

#region Get-ADDomain
function Get-ADDomain
{
	[cmdletBinding()]
	param(		
		[Parameter(Mandatory=$true)]
		[ValidateScript({
			if (-not $_.Contains("."))
			{
				throw "The Name must be a FQDN"
			}
			return $true
		})]
		[string[]]$Name
	)
	
	process
	{
		foreach ($item in $Name)
		{
			$ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $item)
		
			[System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
		}
	}
}
#endregion Get-ADDomain

#---------------------------------------------------

Add-Type -TypeDefinition $typesDefinitions -Language CSharpVersion3 -ReferencedAssemblies System.DirectoryServices
Write-Host "AD Replication Module loaded" -ForegroundColor DarkGreen