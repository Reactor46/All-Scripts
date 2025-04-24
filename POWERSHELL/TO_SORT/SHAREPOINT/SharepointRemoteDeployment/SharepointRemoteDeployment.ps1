#=======================================
# Parameter definition
#=======================================
Param(
[string] $Server,
[string] $IdentityName,
[string] $WebApplication,
[string] $Location,
[System.Management.Automation.PSCredential] $Credentials,
[switch] $InstallSolution,
[switch] $RemoveSolution
)

#==============================================
# Function that validates the script parameters
#==============================================
function ValidateParams
{
	$validInputs = $true
	$errorString =  ""

	if ($Server -eq "")
	{
		$validInputs = $false
		$errorString += "`n`nMissing Parameter: The -Server parameter is required. Please pass in the name of a valid SP Farm Server." + "`n"
	}

	if ($IdentityName -eq "")
	{
		$validInputs = $false
		$errorString += "`nMissing Parameter: The -Identity parameter is required. Please pass in the desired SP Solution Name." + "`n"
	}

	if($InstallSolution -and $RemoveSolution)
	{
		$validInputs = $false
		$errorString += "`nConflicting Parameter: You cannot use the -InstallSolution and -RemoveSolution switches together." + "`n"
	}
	
	if($InstallSolution -and ($Location -eq ""))
	{
		$validInputs = $false
		$errorString += "`nMissing Parameter: You cannot use the -InstallSolution switch without defining a solution to deploy with -Location parameter."
	}

	if (!$validInputs)
	{
		Write-error "$errorString"
	}

	return $validInputs
}

#==========================================================================
# Function that returns true if the incoming argument is a help request
#==========================================================================
function IsHelpRequest
{
	param($argument)
	return ($argument -eq "-?" -or $argument -eq "-help");
}

#===================================================================
# Function that displays the help related to this script following
# the same format provided by get-help or <cmdletcall> -?
#===================================================================
function Usage
{
@"
NAME: SharePointRemoting.ps1

SYNOPSIS:
Executes SharePoint deployment commands on a remote machine

SYNTAX:
SharePointRemoting.ps1
`t[-Server <SPFarm>]
`t[-Identity <SPSolutionName>]
`t[-WebApplication <WebApp>]
`t[-Credentials <Credentials>]
`t[-InstallSolution]
`t[-InstallFeature]
`t[-UninstallFeature]
`t[-RemoveSolution]		

PARAMETERS:
-Server (required)
The SharePoint farm server to execute the remote commands.

-Identity (required)
The name of the SharePoint solution.


-------------------------- EXAMPLE 1 --------------------------



-------------------------- EXAMPLE 2 --------------------------
"@
}

#=======================================
# Check for Usage Statement Request
#=======================================
$args | foreach { if (IsHelpRequest $_) { Usage; exit; } }

#=====================================================
# Validate the parameters and Execute the functions
#=====================================================
$ifValidParams = ValidateParams;

if (!$ifValidParams) { exit; }


function Wait-SPDeployment($solution, [bool]$deploying, [string]$status, [int]$percentComplete) {
do { 
  Start-Sleep 2
  Write-Progress -Activity "Deploying solution $($solution.Name)" -Status $status -PercentComplete $percentComplete
  $solution = Get-SPSolution $solution
  if ($solution.LastOperationResult -like "*Failed*") { throw "An error occurred during the solution retraction, deployment, or update." }
  if (!$solution.JobExists -and (($deploying -and $solution.Deployed) -or (!$deploying -and !$solution.Deployed))) { break }
} while ($true)
sleep 5  
}


#=======================================
# Assign values to Variables
#=======================================

$deploy =
{
	function Wait-SPDeployment($solution, [bool]$deploying, [string]$status, [int]$percentComplete) {
	do { 
	  Start-Sleep 2
	  Write-Progress -Activity "Deploying solution $($solution.Name)" -Status $status -PercentComplete $percentComplete
	  $solution = Get-SPSolution $solution
	  if ($solution.LastOperationResult -like "*Failed*") { throw "An error occurred during the solution retraction, deployment, or update." }
	  if (!$solution.JobExists -and (($deploying -and $solution.Deployed) -or (!$deploying -and !$solution.Deployed))) { break }
	} while ($true)
	sleep 5  
	}

	$IdentityName = $args[0]
	$Location = $args[1]
	$WebApplication = $args[2]
	
	$ver = $host | select version
	if ($ver.Version.Major -gt 1)  {$Host.Runspace.ThreadOptions = "ReuseThread"}
	Add-PsSnapin Microsoft.SharePoint.PowerShell
	Set-location $home
	
    # Get the current solution (if exists
	$solution = Get-SPSolution $IdentityName -ErrorAction SilentlyContinue
    
    if ($solution -ne $null) {
      #Retract the solution
      if ($solution.Deployed) {
        Write-Progress -Activity "Deploying solution $IdentityName" -Status "Retracting $IdentityName" -PercentComplete 0
        if ($solution.ContainsWebApplicationResource) {
          $solution | Uninstall-SPSolution –Identity $IdentityName -AllWebApplications -Confirm:$false
        } else {
          $solution | Uninstall-SPSolution –Identity $IdentityName -Confirm:$false
        }
        #Block until we're sure the solution is no longer deployed.
        Wait-SPDeployment $solution $false "Retracting $IdentityName" 12
        Write-Progress -Activity "Deploying solution $IdentityName" -Status "Solution retracted" -PercentComplete 25
      }

      #Delete the solution
      Write-Progress -Activity "Deploying solution $IdentityName" -Status "Removing $IdentityName" -PercentComplete 30
      Remove-SPSolution –Identity $IdentityName -Confirm:$false
      Write-Progress -Activity "Deploying solution $IdentityName" -Status "Solution removed" -PercentComplete 50
    }


	
    #Add the solution
    Write-Progress -Activity "Deploying solution $IdentityName" -Status "Adding $IdentityName" -PercentComplete 50
    $solution = Add-SPSolution $Location –Confirm:$false
    Write-Progress -Activity "Deploying solution $IdentityName" -Status "Solution Added" -PercentComplete 75
	
	#Install the Solution
    Write-Progress -Activity "Deploying solution $IdentityName" -Status "Installing $IdentityName to all Web Applications" -PercentComplete 75
    Install-SPSolution -Identity $IdentityName -CASPolicies -AllWebApplications -Confirm:$false
    Wait-SPDeployment $solution $true "Installing $IdentityName to all Web Applications" 85

}

$retract =
{
	$IdentityName = $args[0]
	
	function Wait-SPDeployment($solution, [bool]$deploying, [string]$status, [int]$percentComplete) {
	do { 
	  Start-Sleep 2
	  Write-Progress -Activity "Uninstalling solution $($solution.Name)" -Status $status -PercentComplete $percentComplete
	  $solution = Get-SPSolution $solution
	  if ($solution.LastOperationResult -like "*Failed*") { throw "An error occurred during the solution retraction, deployment, or update." }
	  if (!$solution.JobExists -and (($deploying -and $solution.Deployed) -or (!$deploying -and !$solution.Deployed))) { break }
	} while ($true)
	sleep 5  
	}
	
	$ver = $host | select version
	if ($ver.Version.Major -gt 1)  {$Host.Runspace.ThreadOptions = "ReuseThread"}
	Add-PsSnapin Microsoft.SharePoint.PowerShell
	Set-location $home
	
	# Get the current solution (if exists
	$solution = Get-SPSolution $IdentityName -ErrorAction SilentlyContinue
    
    if ($solution -ne $null) {
      #Retract the solution
      if ($solution.Deployed) {
        Write-Progress -Activity "Uninstalling solution $IdentityName" -Status "Retracting $IdentityName" -PercentComplete 0

        Uninstall-SPSolution –Identity $IdentityName -AllWebApplications -Confirm:$false

        #Block until we're sure the solution is no longer deployed.
        Wait-SPDeployment $solution $false "Uninstalling $IdentityName" 12
        Write-Progress -Activity "Uninstalling solution $IdentityName" -Status "Solution retracted" -PercentComplete 25
      }

      #Delete the solution
      Write-Progress -Activity "Uninstalling solution $IdentityName" -Status "Removing $IdentityName" -PercentComplete 30
      Remove-SPSolution –Identity $IdentityName -Confirm:$false
      Write-Progress -Activity "Uninstalling solution $IdentityName" -Status "Solution removed" -PercentComplete 50
    }
}

#=======================================
# Execute functions
#=======================================
#$Server |%{ Invoke-Command -ComputerName $_ -Credential $Credentials -Authentication Credssp -ScriptBlock $script }


if($InstallSolution)
{
	$Server |%{ Invoke-Command -ComputerName $_ -Authentication Kerberos -ScriptBlock $deploy -ArgumentList $IdentityName, $Location, $WebApplication }
}
elseif($RemoveSolution)
{
	$Server |%{ Invoke-Command -ComputerName $_  -Authentication Kerberos -ScriptBlock $retract -ArgumentList $IdentityName }
}