function Get-SnmpTrap {
<#
.SYNOPSIS
Function that will list SNMP Community string, Security Options and Trap Configuration for SNMP version 1 and version 2c.
.DESCRIPTION
** This function will list SNMP settings of  windows server by reading the registry keys under HKLM\SYSTEM\CurrentControlSet\services\SNMP\Parameters\
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