<#

.SYNOPSIS
                This script automates the upgrade of UCS firmware
                
.DESCRIPTION
                This script logs into Cisco.com, download firmware bundles to a local working directory, upload bundles to the target UCS domain and walk through all steps of the upgrade

.EXAMPLE
                UCSMFirmwareUpdate.ps1 -ucs xx.xx.xx.xx -version 'x.x(xx)' -imagedir c:\work\images -hostFirmwarePackDn "
                Upgrades all components of UCS
                -ucs -- UCS Manager IP -- Example: "1.2.3.4"
                -version -- UCS Manager version to upgrade 
                -imagedir -- Path to download firmware bundle
			    All parameters are mandatory
                The only prompts that will always be presented to the user will be for Username and Password


.EXAMPLE
                UCSMFirmwareUpdate.ps1 -ucs xx.xx.xx.xx -version 'x.x(xx)' -imagedir c:\work\images -infraOnly 
                Upgrade Infrastructure only portion of UCSM
                -infraOnly -- Optional SwitchParameter when specified upgrades Infrastructure only
                 
.NOTES
                Author: Nitin Veda
                Email: nveda@cisco.com
                Company: Cisco Systems, Inc.
                Version: v1.1
                Date: 07/15/2014
                Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
                UCSM IP Address
                UCS Manager version to upgrade
                Directory path to download firmware image bundles
				Host Firmware Pack to upgrade Blade Server and Rack Server
                Switch Parameter to upgrade infrastructure only portion

.OUTPUTS
                None
                
.LINK
                https://communities.cisco.com/docs/DOC-36062

#>


param(
    [parameter(Mandatory=${true})][string]${version},
    [parameter(Mandatory=${true})][string]${ucs},
    [parameter(Mandatory=${true})][string]${imageDir},
	[parameter()][Switch]${infraOnly}
)

if ((Get-Module | where {$_.Name -ilike "CiscoUcsPS"}).Name -ine "CiscoUcsPS")
	{
		Write-Host "Loading Module: Cisco UCS PowerTool Module"
		Write-Host ""
		Import-Module CiscoUcsPs
	}  

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


function Check-UcsState {

	$Error.Clear()
	$output = Get-UcsStatus -ErrorAction SilentlyContinue
	
	if (${Error}) 
	{
		Write-Host "ERROR: Lost UCS connection to UCS Manager Domain: '$($ucs)'"
		Write-Host "     Error equals: ${Error}"
		Write-Host ""
		$output = Disconnect-Ucs -ErrorAction SilentlyContinue
   	    $Error.Clear()
		
		Write-Host  "RETRY: Retrying login to UCS Manager Domain: '$($ucs)' ..."
		${myCon} = Connect-Ucs -Name ${ucs} -Credential ${ucsCred} -ErrorAction SilentlyContinue
		if (${Error}) 
		{
			Write-Host "Error creating a session to UCS Manager Domain: '$($ucs)'"
			Write-Host "     Error equals: ${Error}"
			Write-Host "     Sleeping for 60 seconds ..."
			Write-Host ""
	        start-countdown -seconds 60
        }
    }
}

# Script only supports one UCS Domain update at a time
$output = Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs $false

Try
{
    ${Error}.Clear()
   
    ${versionSplit} = ${version}.Split("()")
    ${versionBundle} = ${versionSplit}[0] + "." + ${versionSplit}[1]
    
	if ((${infraOnly} -eq $false))
	{
		${aSeriesBundle} = "ucs-k9-bundle-infra." + ${versionBundle} + ".A.bin"
		${bSeriesBundle} = "ucs-k9-bundle-b-series." + ${versionBundle} + ".B.bin"
		${cSeriesBundle} = "ucs-k9-bundle-c-series." + ${versionBundle} + ".C.bin"
		
		${bundle} = @()
		${bundle} = @(${aSeriesBundle},${bSeriesBundle},${cSeriesBundle})
		${ccoImageList} = @()
	}
	elseif (${infraOnly} -eq $true)
	{
		${aSeriesBundle} = "ucs-k9-bundle-infra." + ${versionBundle} + ".A.bin"
				
		${bundle} = @()
		${bundle} = @(${aSeriesBundle})
		${ccoImageList} = @()
	}
    
	Write-Host "Starting Firmware download process to local directory: ${imageDir}"
	Write-Host ""
	
    foreach(${eachBundle} in ${bundle})
    {
        ${fileName} = ${imagedir} +  "\" + ${eachBundle}
         if( test-path -Path ${fileName})
         {
              Write-Host "Image File : '${eachBundle}' already exist in local directory: '${imageDir}'"
         }
         else
         {
              ${ccoImageList} += ${eachBundle}
         }
    }
    
    if( ${ccoImageList} -ne ${null})
    {
        Write-Host  "Enter Cisco.com (CCO) Credentials"
        ${ccoCred} = Get-Credential
        foreach(${eachbundle} in ${ccoImageList})
        {
            [array]${ccoImage} += Get-UcsCcoImageList -Credential ${ccoCred} | where { $_.ImageName -match ${eachbundle}}
			Write-Host "Preparing to download UCS Manager version '$($version)' bundle file: '$($eachbundle)'"
        }
        Write-Host  "Downloading UCS Manager version: '$($version)' bundles to local directory: $($imageDir)"
		$output = ${ccoImage} | Get-UcsCcoImage -Path ${imageDir}
    }
	Write-Host "Firmware download process completed to local directory: ${imageDir}"
	Write-Host ""

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
		
    foreach (${image} in ${bundle})
    {
		Write-Host "Checking if image file: '$($image)' is already uploaded to UCS Domain: '$($ucs)'"
		${firmwarePackage} = Get-UcsFirmwarePackage -Name ${image}
        ${deleted} = $false
        if (${firmwarePackage})
        {
           	# Check if all the images within the package are present by looking at presence
            ${deleted} = ${firmwarePackage} | Get-UcsFirmwareDistImage | ? { $_.ImageDeleted -ne ""}
        }
    
		if (${deleted} -or !${firmwarePackage})
        {
            # If Image does not exist on FI, uplaod
            $fileName = ${imageDir} +  "\" + ${image}
			Write-Host "Uploading image file: '$($image)' to UCS Domain: '$($ucs)'"
            $output = Send-UcsFirmware -LiteralPath $fileName | Watch-Ucs -Property TransferState -SuccessValue downloaded -PollSec 30 -TimeoutSec 600
        	Write-Host "Upload of image file: '$($image)' to UCS Domain: '$($ucs)' completed"
			Write-Host ""  
		}
		else
		{
			Write-Host "Image file: '$($image)' is already uploaded to UCS Domain: '$($ucs)'"
			Write-Host ""  
		}
    }
          
 

	Write-Host ""
    # Activate UCSM, IOM and FI
    ${firmwareRunningUcsm} = Get-UcsMgmtController -Subject system | Get-UcsFirmwareRunning
    if (${firmwareRunningUcsm}.version -eq ${version})
    {
        Write-Host "UCS Manager 'running' software version already at version: '${version}' on UCS Domain: '$($ucs)'"
    }
    else
    {
        Write-Host "Activating UCS Manager software to version: '${version}' on UCS Domain: '$($ucs)'"
		Write-Host "     Requires a re-login to UCS Manager after UCS Manager Upgrade"
		${infraBundle} = ${version} + "A"
		Get-UcsFirmwareInfraPack | Set-UcsFirmwareInfraPack -InfraBundleVersion ${infraBundle}
        Write-Host "Please wait while UCS Manager restarts on UCS Domain: '$($ucs)'"
		Write-host "     Operation may take 3 or more minutes"
        Try
        {
    		Write-Host "Disconnecting session from UCS Manager Domain: '$($ucs)'"
			$output = Disconnect-Ucs
        }
        Catch
        {
           Write-Host "Error disconnecting session from UCS Manager Domain: '$($ucs)'"
        }
        Write-Host "Sleeping for 3 minutes ..."
		Write-Host ""  
        start-countdown -seconds 180
		${count} = 0
		do
        {
			${count}++
			$Error.Clear()
			Write-Host  "Attempt ${count}: Retrying login to UCS Manager Domain: '$($ucs)' ..."
		    ${myCon} = Connect-Ucs -Name ${ucs} -Credential ${ucsCred} -ErrorAction SilentlyContinue
            
			if (${Error}) 
			{
				Write-Host "Error creating a session to UCS Manager Domain: '$($ucs)'"
				Write-Host "     Error equals: ${Error}"
				Write-Host "     Sleeping for 60 seconds ..."
				Write-Host ""
		        start-countdown -seconds 60
            }
        } while (${myCon} -eq ${null})
	
		Write-Host "Successfully logged back into UCS Domain: '${ucs}'"
	
		${firmwareRunningUcsm} = Get-UcsMgmtController -Subject system | Get-UcsFirmwareRunning
    	if (${firmwareRunningUcsm}.version -eq ${version})
    	{
        	Write-Host "UCS Manager 'running' software version updated to version: '${version}' successfully on UCS Domain: '$($ucs)'"
			Write-Host "Acknowledge Activation of Primary FI on UCS Domain: '$($ucs)'"
    	} 
		else
		{
			Write-Host "UCS Manager 'running' software version 'NOT' updated to version: '${version}' successfully on UCS Domain: '$($ucs)'"
			Write-Host "Exiting"
			exit
		}
    }
	
	# Activate Blade
	if ((${infraOnly} -eq $false))
	{
	 # Write-Host "Upgrading Blade Firmware to version: '$($version)' on UCS Domain: '$($ucs)'"
		${bladeBundle} = ${version} + "B"
		${rackBundle} = ${version} + "C"
		$blades = Get-UcsBlade
        
		foreach (${blade} in ${blades})
		{
			${firmwareRunningBlade} = ${blade} | Get-UcsMgmtController -Subject blade | Get-UcsFirmwareRunning -Deployment system
			if (${firmwareRunningBlade}.version -eq ${version})
			{ 
			    ${bladeDn} = ${blade}.dn
				Write-Host "Blade ('${bladeDn}') running software version already at version: '${version}''"
			}
			else
			{
				${assignToDn} = ${blade}.AssignedToDn
				#if( ${assignToDn} -ne ${null})
				if([string]::IsNullOrEmpty($assignToDn))
				{
					${hostFirmwarePackDn} = "org-root/fw-host-pack-default"
				}
				else
				{
					${spName} = ${assignToDn}.Substring(${assignToDn}.LastIndexOf('/')+4)
					${fwPolicyName} = Get-UcsServiceProfile -Name ${spName}
					${hostFirmwarePackDn} = ${fwPolicyName}.OperHostFwPolicyName
				}
				Get-UcsFirmwareComputeHostPack -Dn ${hostFirmwarePackDn}| Set-UcsFirmwareComputeHostPack -BladeBundleVersion ${bladeBundle}
				
				[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
				$Output= [System.Windows.Forms.MessageBox]::Show("The update process will need to reboot the server(s). Would you like to acknowledge the same?" , "Status" , 4)

				if ($Output -eq "YES" )
				{
					Get-UcsServiceProfile -Type 'instance' -AssocState 'associated'| %{If((-not [string]::IsNullOrEmpty($_.OperHostFwPolicyName)) -and ($_.OperHostFwPolicyName -eq ${hostFirmwarePackDn}))
						{
							Get-UcsLsmaintAck -dn ($_.Dn + '/ack') | Set-UcsLsmaintAck -AdminState 'trigger-immediate'
						}
					}
				}
				else
				{
					Write-Host "Acknowledgement is required to update blade server."
					exit
				}
			}
		}
	}
	
	#Disconnect from UCS
    Write-Host "Install-All update process to version: ${version} executed successfully"
	Write-Host "     Disconnecting from UCS Domain: '${ucs}'"
    Disconnect-Ucs
}
Catch
{
	 Write-Host "Error occurred in script:"
     Write-Host ${Error}
     exit
}