param(
    [parameter(Mandatory=${true})][String]$Ucs,
    [parameter(Mandatory=${true})][String]$UcsOrg,
    [parameter(Mandatory=${true})][String]$UcsSpTemplate,
    [parameter(Mandatory=${true})][String]$UcsBladeDn,
    [parameter(Mandatory=${true})][string]$Hostname
)
                                                                                                                                                                                                 

# Global Variables
$DhcpServer = "mms-scvmm.ucsdemo.cisco.com"
$DhcpScope = "192.168.22.0"               
$DnsServer = "mms-ad.ucsdemo.cisco.com"
$DnsZone = "ucsdemo.cisco.com"
$wdsserver = "mms-scvmm.ucsdemo.cisco.com"
$WdsPxePromptPolicy = "NoPrompt"
$WdsBootProgram = "Boot\x64\Images\pxeboot.com"
$WdsBootImagePath = "Boot\x64\Images\boot.wim"
$WdsClientUnattend = "WdsClientUnattend\bootunattendwin2012.xml"
$WdsDomainUser= "UCSDEMO\runasacct"
$WdsJoinRights = "Full"
$WdsOu = "OU=CiscoUCS,DC=ucsdemo,DC=cisco,DC=com"

function Start-Countdown{

	Param(
		[INT]$Seconds = (Read-Host "Enter seconds to countdown from")
	)

	while ($seconds -ge 1){
	    Write-Progress -Activity "Sleep Timer Countdown" -SecondsRemaining $Seconds -Status "Time Remaining"
	    Start-Sleep -Seconds 1
	$Seconds --
	}
}

# Import Modules
if ((Get-Module |where {$_.Name -ilike "CiscoUcsPS"}).Name -ine "CiscoUcsPS")
	{
	Write-Host "Loading Module: Cisco UCS PowerTool Module"
	Import-Module CiscoUcsPs
	}

if ((Get-Module |where {$_.Name -ilike "DNSServer"}).Name -ine "DNSServer")
	{
	Write-Host "Loading Module: Microsoft DNSServer Module"
	Import-Module DNSServer
	}

if ((Get-Module |where {$_.Name -ilike "DHCPServer"}).Name -ine "DHCPServer")
	{
	Write-Host "Loading Module: Microsoft DHCPServer Module"
	Import-Module DHCPServer
	}

Try
{
    # Get UCS PS Connection
    Write-Host "UCS: Checking for current UCS Connection for UCS Domain: '$($ucs)'"
    $UcsConn = Get-UcsPSSession | where { $_.Name -eq $Ucs }
    if ( ($UcsConn).Name -ne $Ucs ) {
        Write-Host "UCS: UCS Connection: '$($ucs)' is not connected"
        Write-Host "UCS: Enter UCS Credentials"
        $cred = Get-Credential
        $UcsConn = connect-ucs $Ucs -Credential $cred
    }

    # Get UCS org in connected UCS session
    Write-Host "UCS: Checking for UCS Org: '$($UcsOrg)' on UCS Domain: '$($ucs)'"
    $TargetOrg = Get-UcsOrg -Name $UcsOrg
    if ( $TargetOrg -eq $null ) {
        Write-Host "UCS: UCS Organization: '$($TargetOrg)' is not present"
        exit
    }

    # Get UCS SP template on connected UCS session
    Write-Host "UCS: Checking for UCS Service Profile: '$($UcsSpTemplate)' on UCS Domain: '$($ucs)'"
    $TargetSpTemplate = $TargetOrg | Get-UcsServiceProfile -Name $UcsSpTemplate -ucs $UcsConn -LimitScope
    if ( $TargetSpTemplate -eq $null ) {
        Write-Host "UCS: UCS Service Profile Template: '$($TargetSpTemplate.Dn)' is not present"
        exit
    } elseif ( ($TargetSpTemplate).Type -notlike "*-template*" ) {
        Write-Host "UCS: UCS Service Profile: '$($TargetSpTemplate.Dn)' is not a service profile template"
        exit
    }   

    # Get UCS Blade on connected UCS session, check availability of UCS Blade
    Write-Host "UCS: Checking availability on UCS Blade: '$($UcsBladeDn)' on UCS Domain: '$($ucs)'"
    $TargetBlade = Get-UcsBlade -dn $UcsBladeDn 
    if ( $TargetBlade -eq $null ) {
        Write-Host "UCS: UCS Blade: '$($TargetBlade.Dn)' is not present"
        exit
    } elseif ( ($TargetBlade).Association -ne "none" -and ($TargetBlade).Availability -ne "available" ) {
        Write-Host "UCS: UCS Blade: '$($TargetBlade.Dn)' is not available"
        exit
    }

    # Check to see if SP is already created on connected UCS Session
    Write-Host "UCS: Checking to see if SP: '$($hostname)' exists on UCS Domain: '$($ucs)'"
    $SpToCreate = $TargetOrg | Get-UcsServiceProfile -Name $Hostname -LimitScope
    if ( $SpToCreate -ne $null ) {
        Write-Host "UCS: UCS Service Profile: '$($Hostname)' is already created"
        exit
    }

	# Create New UCS SP from Template
    Write-host "UCS: Creating new SP: '$($hostname)' from UCS SP Template: '$($TargetSpTemplate.Dn)' on UCS Domain: '$($ucs)'"
    $NewSp = Add-UcsServiceProfile -org $TargetOrg -Ucs $UcsConn -SrcTemplName ($TargetSpTemplate).Name -Name $Hostname 
    $devnull = $NewSp | Set-UcsServerPower -ucs $UcsConn -State "down" -Force

    # Associate Service Profile to Blade
   	Write-Host "UCS: Associating new UCS SP: '$($NewSp.Name)' to UCS Blade: '$($TargetBlade.Dn)' on UCS Domain: '$($Ucs)'"
    $devnull = Associate-UcsServiceProfile -ucs $UcsConn -ServiceProfile $NewSp -Blade $TargetBlade -Force


	# Monitor UCS SP Associate for completion
	Write-Host "UCS: Waiting for UCS SP: '$($NewSp.name)' to complete SP association process on UCS Domain: '$($Ucs)'"
    Write-host "Sleeping 3 minutes"
    Start-Countdown -seconds 180

    $i = 0

		do {
			if ( (Get-UcsServiceProfile -Dn $NewSp.Dn).AssocState -ieq "associated" )
			{
				break
			} else {
                Write-host "Sleeping 30 seconds"
    			Start-Countdown -seconds 30
                $i++
                Write-Host "UCS: RETRY $($i): Checking for UCS SP: '$($NewSp.name)' to complete SP association process on UCS Domain: '$($Ucs)'"
                
            }		
        } until ( (Get-UcsServiceProfile -Dn $NewSp.Dn).AssocState -ieq "associated" -or $i -eq 24 )

    if ( $i -eq 24 ) {
    	Write-Host "UCS: Association process of UCS SP: '$($NewSp.name)' failed on UCS Domain: '$($Ucs)'"	
        exit
    } 
    
   	Write-Host "UCS: Association process of UCS SP: '$($NewSp.name)' completed on UCS Domain: '$($Ucs)'"	

    # Get DHCP Scope from DHCP Server
    Write-Host "DHCP: Checking for DHCP Scope: '$($DhcpScope)' on DHCP Server: '$($DhcpServer)'"
    $TargetDhcpScope = Get-DhcpServerv4Scope -ScopeId $DhcpScope -ComputerName $DhcpServer
    if ( $TargetDhcpScope -eq $null ) {
        Write-Host "DHCP: DHCP Scope: '$($TargetDhcpScope)' is not present"
        exit
    }
    
    # Check for available IP in DHCP Scope
    Write-Host "DHCP: Checking for Available IP in DHCP Scope: '$($DhcpScope)' on DHCP Server: '$($DhcpServer)'"
    $AvailableIP = $TargetDhcpScope | Get-DhcpServerv4FreeIPAddress
    if ( $AvailableIP -eq $null ) {
        Write-Host "DHCP: No Available IPs in DHCP Scope: '$($TargetDhcpScope)' are present"
        exit
    } else {
        Write-Host "DHCP: IP Address: '$($AvailableIP)' in DHCP Scope: '$($TargetDhcpScope)' is available"
    }
    
    # Create DHCP Lease for IP / Hostname / SP
    Write-host "UCS: Getting MAC address for new SP: '$($hostname)' for vNIC: 'eth0' on UCS Domain: '$($ucs)'"
    $SpMacAddr = ($NewSp | Get-UcsVnic -Name "eth0").Addr 
    if ( $SpMacAddr -eq $null ) {
        Write-Host "UCS: No vNIC named 'eth0' found on SP: '$($hostname)' on UCS Domain: '$($ucs)'"
        exit
    } else {
        Write-Host "UCS: MAC Address: '$($SpMacAddr)' found on vNIC named 'eth0' found in SP: '$($hostname)' on UCS Domain: '$($ucs)'"
    }
    
    Write-Host "DHCP: Adding DHCP Reservation for IP: '$($AvailableIP)' with MAC Address: '$($SpMacAddr)' and Hostname: '$($Hostname)' in DHCP Scope: '$($TargetDhcpScope)' is available"

    $SpMacAddr = $SpMacAddr -replace ':',''
    $devnull = Add-DhcpServerv4Reservation -ComputerName $DhcpServer -ScopeId ($TargetDhcpScope).ScopeId -IPAddress $AvailableIP -ClientId $SpMacAddr -Name $Hostname -Description "Added automatically" -Type Dhcp 
   
    # Get DNS Zone from DNS Server
    Write-Host "DNS: Checking for DNS Zone: '$($DnsZone)' on DNS Server: '$($DnsServer)'"
    $TargetDnsZone = Get-DnsServerZone -Name $DnsZone -ComputerName $DnsServer
    if ( $TargetDnsZone -eq $null ) {
        Write-Host "DNS: DNS Zone: '$($TargetDnsZone)' is not present"
        exit
    }

    Write-Host "DNS: Adding DNS A Record for IP: '$($AvailableIP)' and Hostname: '$($Hostname)' in DNS Zone: '$($TargetDnsZone)'"
    $truncatedhost = $Hostname -replace ".$($DnsZone)",''
    $devnull = $TargetDnsZone | Add-DnsServerResourceRecordA -Name $truncatedhost -IPv4Address $AvailableIP -ComputerName $DnsServer 

    Write-Host "WDS: Adding an AD Pre-staged Device Record for Hostname: '$($Hostname)' on WDS Server: $($WdsServer)"
    $devnull = WDSUTIL.exe /Add-Device /Force /Device:$truncatedhost /ID:$SpMacAddr /Group:CiscoUCS /ReferralServer:$WdsServer /WDSClientUnattend:$WdsClientUnattend /User:$WdsUser /JoinRights:$WdsJoinRights /JoinDomain:Yes /OU:$WdsOu /PxePromptPolicy:$WdsPxePromptPolicy /BootImagePath:$WdsBootImagePath /BootProgram:$WdsBootProgram

    # Set SP Power State to Up		
	Write-Host "UCS: Setting Desired Power State to 'up' of Service Profile: '$($NewSp.name)' on UCS Domain: '$($Ucs)'"
	$PowerSpOn = $NewSp | Set-UcsServerPower -ucs $UcsConn -State "up" -Force
}
Catch
{
	 Write-Host "Error occurred in script:"
     Write-Host ${Error}
     exit
}