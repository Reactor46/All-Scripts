param(
[string]$DeploymentPath
)

Add-PSSnapin Microsoft.SharePoint.Powershell

function Install-Solution([string]$path, [bool]$gac, [bool]$cas, [string[]]$webApps = @())
{
    $spAdminServiceName = "SPAdminV4"

    [string]$name = Split-Path -Path $path -Leaf
    $solution = Get-SPSolution $name -ErrorAction SilentlyContinue

    if ($solution -ne $null) {
        #Retract the solution
        if ($solution.Deployed) {
            Write-Host "Retracting solution $name..."
            if ($solution.ContainsWebApplicationResource) {
                $solution | Uninstall-SPSolution -AllWebApplications -Confirm:$false
            } else {
                $solution | Uninstall-SPSolution -Confirm:$false
            }
            Stop-Service -Name $spAdminServiceName
            Start-SPAdminJob -Verbose
            Start-Service -Name $spAdminServiceName

            #Block until we're sure the solution is no longer deployed.
            do { Start-Sleep 2 } while ((Get-SPSolution $name).Deployed)
        }

        #Delete the solution
        Write-Host "Removing solution $name..."
        Get-SPSolution $name | Remove-SPSolution -Confirm:$false
    }

    #Add the solution
    Write-Host "Adding solution $name..."
    $solution = Add-SPSolution $path

    #Deploy the solution
    if (!$solution.ContainsWebApplicationResource) {
        Write-Host "Deploying solution $name to the Farm..."
        $solution | Install-SPSolution -GACDeployment:$gac -CASPolicies:$cas -Confirm:$false -Force
    } else {
        if ($webApps -eq $null -or $webApps.Length -eq 0) {
            Write-Warning "The solution $name contains web application resources but no web applications were specified to deploy to."
            return
        }
        $webApps | ForEach-Object {
            Write-Host "Deploying solution $name to $_..."
            $solution | Install-SPSolution -GACDeployment:$gac -CASPolicies:$cas -WebApplication $_ -Confirm:$false -Force
        }
    }
    Stop-Service -Name $spAdminServiceName
    Start-SPAdminJob -Verbose
    Start-Service -Name $spAdminServiceName

    #Block until we're sure the solution is deployed.
    do { Start-Sleep 2 } while (!((Get-SPSolution $name).Deployed))
}

#Setup Path Variables
$RootPath = "d:\SharePoint\Deployment"
#$DeployedPath = $RootPath + "\Deployed"
$InstallPath = $RootPath + "\Install"
#$UpgradePath = $RootPath + "\Upgrade"

#Get the web app name and version for the deployment from the user
#$DeploymentPath = read-host -prompt "What is the name of the target web application (ie...'http://authoring.kelsey-seybold.com')"
#$DeploymentPath = $args[0]
Write-Host $DeploymentPath

$StartIndex = $DeploymentPath.IndexOf("//")
#$DeploymentFolderName = $DeploymentPath.Substring($StartIndex+2)

#$HasPortInFolder = $DeploymentFolderName.IndexOf(":")

#if ($HasPortInFolder -gt 0)
#{
    #remove colon from path, folder names cannot contain colons
#    $DeploymentFolderName = $DeploymentFolderName -replace '[:]'
#}

#$SolutionDeployedRootPath = $DeployedPath + "\" + $DeploymentFolderName

#$Version = read-host -prompt "What is the version number for the deployment of the selected SharePoint solutions? (ie...'V1.0')"
#$SolutionVersionDeployedPath = $SolutionDeployedRootPath + "\" + $Version

#if this is the first deployment to this web app then create the archive directory
#if (!(Test-Path -path $SolutionDeployedRootPath ))
#{
#	New-Item $SolutionDeployedRootPath -type directory
#}

#create the version folder if it doesn't exist
#if (!(Test-Path -path $SolutionVersionDeployedPath ))
#{
#	New-Item $SolutionVersionDeployedPath -type directory
#}

#Loop through all solution files and install them into the farm
$InstallPath
$fileEntries = [IO.Directory]::GetFiles($InstallPath); 
foreach($fileName in $fileEntries) 
{ 
     	Install-Solution $fileName $true $false @($DeploymentPath)
	    Remove-Item $fileName
}  




