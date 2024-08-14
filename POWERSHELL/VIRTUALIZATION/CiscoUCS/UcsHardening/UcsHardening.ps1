<#

.SYNOPSIS
                This script help users secure Cisco Unified Computing System (Cisco UCS) platform devices to improve network security
                
.DESCRIPTION
                This script connects to UCS system and give recommendation around management, control, and data planes in a UCS system.
.EXAMPLE
                UCSHardening.ps1 -ucs xx.xx.xx.xx 
			    All parameters are mandatory
                The only prompts that will always be presented to the user will be for Username and Password

.EXAMPLE
                UCSHardening.ps1 -ucs xx.xx.xx.xx
				Provide recommendation to improve network security
                 
.NOTES
                Author: Nitin Veda
                Email: nveda@cisco.com
                Company: Cisco Systems, Inc.
                Version: v1.1
                Date: 01/23/2015
                Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
                UCSM IP Address
             
.OUTPUTS
                None
                
#>


param(
    [parameter(Mandatory=${true})][string]${ucs}
)

if ((Get-Module | where {$_.Name -ilike "CiscoUcsPS"}).Name -ine "CiscoUcsPS")
	{
		Write-Host "Loading Module: Cisco UCS PowerTool Module"
		Write-Host ""
		Import-Module CiscoUcsPs
	}  

# Script only supports one UCS Domain update at a time
$output = Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs $false

Try
{
    ${Error}.Clear()
   
	# Login into UCS
	Write-Host  "Enter UCS Credentials of UCS Manager to be upgraded to version: '$($version)'"
	${ucsCred} = Get-Credential
	Write-Host ""
	
	Write-Host "Logging into UCS Domain: '$($ucs)'"
	Write-Host ""  
    ${myCon} = Connect-Ucs -Name ${ucs} -Credential ${ucsCred} -ErrorAction SilentlyContinue
    
	if (${Error}) 
	{
		Write-Host "Error creating a session to UCS Manager Domain: '$($ucs)'"
		Write-Host "     Error equals: ${Error}"
		Write-Host "     Exiting"
        exit
    }	
		
	# The Password Strength option is used to require strong passwords. Ensure Password Strength Check is enabled and do not disable it.
	${result} = Get-UcsSecurityGlobalConfig
	
	if (${result}.PwdStrengthCheck -eq "no")
    {
		Write-Host "Password Strength option is used to require strong passwords. Ensure Password Strength Check is enabled and do not disable it."
        Write-Host ""
    }
    
    # Disable  telnet, issue the disable telnet-server command to disable the Telnet service on a Cisco UCS device.
    ${getTelnet} = Get-UcsTelnet
	
	if (${getTelnet}.AdminState  -eq "enable")
    {
		Write-Host "Issue the disable telnet-server command to disable the Telnet service on a Cisco UCS device."
        Write-Host ""
    }
     
	
	# Users not actively administrating should have their Account Status in a status inactive. Accounts can be set to expire at certain time intervals using # the Expire Account Timeframe configuration option.
	
	${userCount} = Get-UcsLocalUser 
	${activeUser} = Get-UcsLocalUser | where {$_.AccountStatus -eq "active"}
	
	Write-Host 'There is currently '${userCount}.Count' users. Out of this '${activeUser}.Count' are currently active. Accounts can be set to expire at certain time intervals using the Expire Account Timeframe configuration option if NOT in use.'
	Write-Host ""
	
	# Fault conditions are logged, cleared, or stored for a configurable interval. When the retention interval field is set, then this configures the length # of time the system retains the fault in memory on the Cisco UCS fabric. It can be forever or a set amount of time.
	
	${faultPolicy} = Get-UcsFaultPolicy 
	
	if (${faultPolicy}.ClearAction -eq "retain")
	{
		Write-Host "Fault conditions are logged, cleared, or stored for a configurable interval. When the retention interval field is set, then this configures the length of time the system retains the fault in memory on the Cisco UCS fabric. It can be forever or a set amount of time."
        Write-Host ""
	}
	
	# Unsecured protocols are disabled by default. These include Telnet and HTTP. HTTP requests will be redirected to HTTPS when HTTPS is enabled, which is #the default setting
	
	${getHttp} = Get-UcsHttp
	${getTelnet} = Get-UcsTelnet
	
	if ((${getHttp}.AdminState -eq "disabled") -and (${getTelnet}.AdminState -eq "disabled"))
	{
		Write-Host "Unsecured protocols are disabled by default. These include Telnet and HTTP. HTTP requests will be redirected to HTTPS when HTTPS is enabled, which is the default setting."
        Write-Host ""
	}
	
	# SNMP is disabled by default. UCS supports SNMP versions 1, 2, and 3. Use SNMPv3
	
	${getSnmp} = Get-UcsSnmp
	
	if (${getSnmp}.AdminState -eq "disabled")
	{
		Write-Host "SNMP is disabled by default. UCS supports SNMP versions 1, 2, and 3. Use SNMPv3."
        Write-Host ""
	}
	
	#Web session limits are configurable for the entire system and per user. The default for the system is 256 and 256 for each user. It is suggested to have #a limit on the number of user sessions (1–2) and a limit for the maximum amount of connections for the total system based on how many users there are.
	
	${sessionLimit} = Get-UcsWebSessionLimit
	
	Write-Host 'Web session limits are configurable for the entire system and per user. The default for the system is 256 and 256 for each user. It is suggested to have a limit on the number of user sessions (1–2) and a limit for the maximum amount of connections for the total system based on how many users there are. Currently sessions per user is set to '${sessionLimit}.SessionsPerUser'.'
    Write-Host ""
	
   	#Disconnect from UCS
  	Write-Host "     Disconnecting from UCS Domain"
    Disconnect-Ucs
}
Catch
{
	 Write-Host "Error occurred in script"
     Write-Host ${Error}
     exit
}