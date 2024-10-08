<#
.SYNOPSIS 
UCSM Update Service Profile Host Firmware Policy

.DESCRIPTION
Used to update the Host Firmware Policy to a more recent version

.PARAMETER UCSMIP
IP Address or comma seperated set if IP Addresses of the UCSM instance/s.  If an IP is entered here the UCSM .csv list will be
ignored.

.PARAMETER UCSMCSV
Full path to a .csv formatted list of UCSM IP's to interact with.  Must be 2 rows with a "UCSM" header in row 1, and a "IP" header in row 2.

.PARAMETER FIRMWARETARGETPOLICY
Target Host Firmware Policy.  You will need to create this via UCSM, Under HostFirmwarePolicy.

.PARAMETER ORGTARGET
Target Orginization you want to update.  Default is root.


.INPUTS
None.

.OUTPUTS
None.

.EXAMPLE
Following will run a validation pass against a single UCSM instance.
.\Set-SPFirmware -UCSMIP 10.10.10.10 -FirmwareTargetPolicy "1.43s"

.EXAMPLE
Following will run a validation test against every Fabric Interconnect in the .csv file.
.\Set-SPFirmware -UCSMCSV "C:\temp\MyUCSMList.csv" -FirmwareTargetPolicy "1.43s"

.EXAMPLE
Following will target a specific organization containing Service Profiles.
.\Set-SPFirmware -UCSMIP 10.10.10.10 -FirmwareTargetPolicy "1.43s" -orgtarget root/Core-Testing

#>
param(
      [string]$UCSMIP,
      [string]$UCSMCSV,
	  [string]$ORGTARGET,
	  [string]$FIRMWARETARGETPOLICY
)
#_______________________________________________________________________________
#__________________ GLOBALS_____________________________________________________
#_______________________________________________________________________________
$outputDir = "C:\SPUpdateOut.txt"
$failedLog = "c:\SPFailedUpdate.txt"
write get-Date > $outputLog
write get-Date > $failedLog
$ReportErrorShowExceptionClass = $true
[array]$spWorking = $null
#_______________________________________________________________________________
#__________________ FUNCTIONS __________________________________________________
#_______________________________________________________________________________
Function Get-UCSMList 
{
	if (!($UCSMIP) -and !($UCSMCXV))
	{
		Write-Host "Error: No UCSM instance specified or CSV of UCSM instances specified."
		Write-Host 'Use "-ucsmip {YourIP}", or "-ucsmcsv {csv path}" to specify a UCSM instance or instances'
		Write-Host "Exiting"
		exit 1
	}
	if ($UCSMIP)
	{
	    [array]$mgrList += $UCSMIP
	}
	  if ($UCSMCSV)
	  {
            if (Get-Item $UCSMCSV -ErrorAction SilentlyContinue)
            {
                  $mgrList = Import-Csv $UCSMCSV
            }
            else
            {
                  Write-Host "Error: CSV File not found at $UCSMCSV."
                  Write-Host "Exiting...."
                  exit 1
            }
      }
      return $mgrList
      
}
Function Validate-FIConnect ([string]$ucsm)
{
      $cred = Get-Credential -Credential "admin"
      Try
      {
            Connect-Ucs $ucsm -Credential $cred
      }
      Catch
      {
			Write "Error Connecting: Most likely the password you entered is incorrect or you have an invalid IP for the UCSM Manager."
            Exit 1
      }
}
Function Get-SPs ([string]$ORGTARGET)
{
    #Write-Host "Validating existing service profiles exist in organization: $($ORGTARGET)"
	try 
	{
		if ($ORGTARGET -eq "")
    	{
	       $targetSPs = Get-UcsServiceProfile -OperState "power-off"
	    }
        else 
        {
           $targetSPs = (Get-UcsOrg  -Dn ${ORGTARGET} | Get-UcsServiceProfile -OperState "power-off")  
        }
	}
	catch 
	{
	[Exception]
		write "There are no service profiles associated with Orgination: $($ORGTARGET).  Can not continue.  Exiting..."
		exit 1
	}
    return $targetSPs
}
Function Validate-FirmwareTarget ([string]$FIRMWARETARGETPOLICY, [string]$ORGTARGET)
{
    if ($ORGTARGET -eq "")
    {
       $hostpacks = @(Get-UcsOrg -Level root |  Get-UcsFirmwareComputeHostPack -Name ${FIRMWARETARGETPOLICY} -LimitScope)
    }
    else
    {
       $hostpacks = @(Get-UcsOrg  -Dn ${ORGTARGET} | Get-UcsFirmwareComputeHostPack -Name ${FIRMWARETARGETPOLICY} -LimitScope)
    }
    if (${hostpacks}.length -lt 1)
    {
        write "host firmware policy does not exist, Exiting.. "
        exit 1
    }
}
Function Update-SPFirmware ([string]$FIRMWARETARGETPOLICY, $sps)
{
    try
	{
       ${sps} | ? { $_.HostFwPolicyName -ne "${FIRMWARETARGETPOLICY}" } | Set-UcsServiceProfile -HostFwPolicyName ${FIRMWARETARGETPOLICY} -Force | Set-UcsServerPower -State up -Force | out-null
    }
	catch
	{
	   write "Reason: Could not set Host Firmware Policy to Service Profile"
	   write "Reason: Could not set Host Firmware Policy to Service Profile" >> $failedLog
	}
}
function Validate-MaintenancePolicy
{
}
Function Set-PowerState 
{
}
#_____________________________________________________________________________
#__________________MAIN PROGRAM ________________________________________________
#_______________________________________________________________________________
$UCSMList = Get-UCSMList
foreach ($ucsm in $ucsmList)
{
	if (!($cred))
	{
	  	$cred = Get-Credential
	}
	$connID = [Guid]::NewGuid().ToString()
    $mycon = Connect-Ucs -Name $ucsm -Credential $cred
	Validate-FirmwareTarget $FIRMWARETARGETPOLICY $ORGTARGET
	$spWorking = @(Get-SPs $ORGTARGET)
    if (${spWorking}.Length -lt 1)
    {
        write "No required service profile(s) found"
    }
    else 
    {
        Update-SPFirmware $FIRMWARETARGETPOLICY $spWorking
    }
    Disconnect-Ucs 
}

